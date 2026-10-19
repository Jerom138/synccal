"""Backend Outlook classique (lecture/ecriture) pilote en COM via pywin32.

Cette voie ne demande ni enregistrement d'application, ni consentement, ni
intervention de l'IT : elle reutilise la session deja authentifiee du client
Outlook installe. En contrepartie, elle exige une session Windows ouverte.
"""

from __future__ import annotations

from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path

from backends.base import BackendError
from core.model import (
    BUSY_BUSY,
    BUSY_ELSEWHERE,
    BUSY_FREE,
    BUSY_OOF,
    BUSY_TENTATIVE,
    MIRROR_CATEGORY,
    PROP_HASH,
    PROP_ORIGIN,
    CanonicalEvent,
    MirrorStyle,
    OriginRef,
    Window,
)

OL_FOLDER_CALENDAR = 9
OL_APPOINTMENT_ITEM = 1
OL_TEXT = 1
OL_NON_MEETING = 0
# olMeetingCanceled / olMeetingReceivedAndCanceled : Outlook conserve la reunion
# dans l'agenda apres annulation, barree et prefixee "Annule:". La recopier
# encombrerait les deux autres agendas sans rien apporter.
OL_CANCELLED_STATUS = (5, 7)

_BUSY_FROM_OL = {
    0: BUSY_FREE,
    1: BUSY_TENTATIVE,
    2: BUSY_BUSY,
    3: BUSY_OOF,
    4: BUSY_ELSEWHERE,
}
_BUSY_TO_OL = {v: k for k, v in _BUSY_FROM_OL.items()}

_FALLBACK_RESTRICT_FMT = "%m/%d/%Y %H:%M"
_LOCALE_USER_DEFAULT = 0x0400
_LOCALE_SSHORTDATE = 0x0000001F


def _com_to_utc(value) -> datetime:
    """Convertit une date COM lue sur StartUTC/EndUTC en datetime UTC.

    On reconstruit a partir des composants plutot que de se fier au tzinfo :
    selon la version de pywin32, les dates COM reviennent naives ou etiquetees
    de facon surprenante, alors que les composants de StartUTC sont, eux,
    toujours l'heure UTC.
    """
    return datetime(
        value.year, value.month, value.day, value.hour, value.minute, value.second,
        tzinfo=timezone.utc,
    )


def _com_local_date(value) -> date:
    """Date locale d'une valeur COM lue sur Start/End (heure murale locale)."""
    return date(value.year, value.month, value.day)


def _restrict_format() -> str:
    """Format de date accepte par Items.Restrict sur ce poste.

    Restrict interprete les dates selon les parametres regionaux de Windows :
    un format US envoye a un Outlook francais serait mal interprete, et la
    fenetre lue serait silencieusement fausse. On derive donc le format du
    format court declare par le systeme.
    """
    try:
        import ctypes

        buffer = ctypes.create_unicode_buffer(128)
        written = ctypes.windll.kernel32.GetLocaleInfoW(
            _LOCALE_USER_DEFAULT, _LOCALE_SSHORTDATE, buffer, 128
        )
        pattern = buffer.value if written else ""
        if not pattern:
            return _FALLBACK_RESTRICT_FMT
    except Exception:  # noqa: BLE001 - pas de locale lisible, on retombe sur l'US
        return _FALLBACK_RESTRICT_FMT

    result = []
    index = 0
    while index < len(pattern):
        char = pattern[index]
        run = 1
        while index + run < len(pattern) and pattern[index + run] == char:
            run += 1
        if char == "d":
            result.append("%A" if run > 3 else ("%a" if run == 3 else "%d"))
        elif char == "M":
            result.append("%B" if run > 3 else ("%b" if run == 3 else "%m"))
        elif char == "y":
            result.append("%Y" if run >= 3 else "%y")
        elif char == "'":
            run = pattern.find("'", index + 1)
            run = (run - index + 1) if run != -1 else 1
            result.append(pattern[index + 1 : index + run - 1])
        else:
            result.append(char * run)
        index += run

    date_fmt = "".join(result)
    if "%" not in date_fmt:
        return _FALLBACK_RESTRICT_FMT
    # Heure en 24 h : accepte aussi bien par un Outlook francais qu'anglophone.
    return date_fmt + " %H:%M"


def _set_utc(appt, field: str, moment: datetime) -> None:
    """Affecte StartUTC ou EndUTC, en UTC explicite.

    Il faut passer par les proprietes UTC, et non par Start/End : pywin32
    convertit tout datetime Python en UTC avant de le remettre a COM, et Outlook
    reinterprete ensuite cette valeur comme une heure locale. Un rendez-vous de
    10:00 a Paris se retrouvait ainsi cree a 08:00, decale de l'offset du fuseau.
    Un datetime naif subit exactement le meme sort : seules StartUTC/EndUTC
    donnent un resultat juste.

    Repli : une chaine ISO en heure locale, que COM interprete correctement
    comme de l'heure murale.
    """
    value = moment.astimezone(timezone.utc).replace(microsecond=0)
    try:
        setattr(appt, field, value)
    except Exception:  # noqa: BLE001
        local = value.astimezone().strftime("%Y-%m-%d %H:%M:%S")
        setattr(appt, field.replace("UTC", ""), local)


def _set_wall_clock(appt, field: str, moment: datetime) -> None:
    """Affecte Start ou End en heure murale locale, pour les journees entieres.

    Une journee entiere est calee sur minuit local : la passer en UTC la
    decalerait d'un jour une fois sur deux. La chaine ISO est le seul format
    non ambigu quels que soient les parametres regionaux.
    """
    setattr(appt, field, moment.replace(microsecond=0).strftime("%Y-%m-%d %H:%M:%S"))


class OutlookComBackend:
    def __init__(self, key: str, label: str, settings: dict, style: MirrorStyle, root: Path) -> None:
        self.key = key
        self.label = label
        self.style = style
        # Un agenda partage par quelqu'un d'autre n'accorde pas toujours le
        # droit de modification. Le declarer en lecture seule evite d'empiler
        # des echecs d'ecriture a chaque run : le moteur n'emettra rien vers lui.
        self.writable = bool(settings.get("writable", True))
        self._store_name = (settings.get("store") or "").strip()
        self._folder_path = (settings.get("folder_path") or "").strip()
        self._shared_with = (settings.get("shared_with") or "").strip()
        self._namespace = None
        self._folder = None
        self._store_id: str | None = None

    # -- connexion ---------------------------------------------------------

    def _connect(self):
        if self._folder is not None:
            return self._folder
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise BackendError(
                self.key + ": pywin32 absent, lancer pip install -r requirements.txt"
            ) from exc

        try:
            pythoncom.CoInitialize()
        except Exception:  # noqa: BLE001 - deja initialise, sans consequence
            pass

        try:
            app = win32com.client.Dispatch("Outlook.Application")
            self._namespace = app.GetNamespace("MAPI")
        except Exception as exc:  # noqa: BLE001
            raise BackendError(
                self.key + ": Outlook classique injoignable (" + str(exc) + "). "
                "Verifier qu'un profil existe et que OUTLOOK.EXE demarre."
            ) from exc

        self._folder = self._resolve_folder()
        return self._folder

    def _resolve_folder(self):
        """Localise le dossier d'agenda, selon trois modes de configuration.

        `shared_with` : l'agenda par defaut d'une autre personne, partage avec
        nous. C'est le mode utile quand un compte ne peut pas etre ajoute au
        profil mais que son proprietaire partage son agenda en ecriture.

        `store` (+ `folder_path` optionnel) : un compte du profil, et
        eventuellement un sous-calendrier plutot que le calendrier par defaut.
        """
        if self._shared_with:
            return self._resolve_shared()

        stores = self._namespace.Stores
        names: list[str] = []
        wanted = self._store_name.strip().casefold()

        for index in range(1, stores.Count + 1):
            store = stores.Item(index)
            try:
                display = store.DisplayName
            except Exception:  # noqa: BLE001 - magasin non monte
                continue
            names.append(display)
            if wanted and display.strip().casefold() != wanted:
                continue
            try:
                folder = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
            except Exception:  # noqa: BLE001 - magasin sans calendrier (PST d'archive)
                continue
            self._store_id = store.StoreID
            return self._descend(folder)

        if self._store_name:
            raise BackendError(
                self.key + ": magasin '" + self._store_name + "' introuvable dans le profil Outlook. "
                "Magasins disponibles : " + (", ".join(names) or "aucun")
            )
        raise BackendError(self.key + ": aucun calendrier Outlook exploitable dans le profil")

    def _resolve_shared(self):
        try:
            recipient = self._namespace.CreateRecipient(self._shared_with)
            recipient.Resolve()
            if not recipient.Resolved:
                raise BackendError(
                    self.key + ": destinataire '" + self._shared_with + "' non resolu. "
                    "Verifier l'adresse, et que son agenda vous est bien partage."
                )
            folder = self._namespace.GetSharedDefaultFolder(recipient, OL_FOLDER_CALENDAR)
        except BackendError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise BackendError(
                self.key + ": agenda partage de '" + self._shared_with + "' inaccessible ("
                + str(exc) + "). Le partage doit inclure le droit de modification."
            ) from exc
        try:
            self._store_id = folder.Store.StoreID
        except Exception:  # noqa: BLE001 - certains agendas partages n'exposent pas de magasin
            self._store_id = None
        return self._descend(folder)

    def _descend(self, folder):
        """Suit `folder_path` sous le calendrier, s'il est configure."""
        if not self._folder_path:
            return folder
        current = folder
        for part in [p for p in self._folder_path.replace("\\", "/").split("/") if p]:
            try:
                current = current.Folders[part]
            except Exception as exc:  # noqa: BLE001
                available = []
                try:
                    available = [current.Folders.Item(i).Name for i in range(1, current.Folders.Count + 1)]
                except Exception:  # noqa: BLE001
                    pass
                raise BackendError(
                    self.key + ": sous-dossier '" + part + "' introuvable sous '" + current.Name
                    + "'. Disponibles : " + (", ".join(available) or "aucun")
                ) from exc
        return current

    def _item_by_id(self, native_id: str):
        self._connect()
        try:
            if self._store_id:
                return self._namespace.GetItemFromID(native_id, self._store_id)
            return self._namespace.GetItemFromID(native_id)
        except Exception:  # noqa: BLE001 - item deja supprime ou deplace
            return None

    # -- lecture -----------------------------------------------------------

    def fetch(self, window: Window) -> list[CanonicalEvent]:
        folder = self._connect()
        try:
            items = folder.Items
            # L'ordre compte : Outlook ignore silencieusement IncludeRecurrences
            # si la collection n'a pas ete triee sur [Start] au prealable.
            items.Sort("[Start]")
            items.IncludeRecurrences = True

            fmt = _restrict_format()
            local_start = window.start.astimezone()
            local_end = window.end.astimezone()
            restriction = (
                "[Start] < '" + local_end.strftime(fmt) + "' AND "
                "[End] > '" + local_start.strftime(fmt) + "'"
            )
            selection = items.Restrict(restriction)
        except Exception as exc:  # noqa: BLE001
            raise BackendError(self.key + ": lecture impossible (" + str(exc) + ")") from exc

        events: list[CanonicalEvent] = []
        try:
            item = selection.GetFirst()
            while item is not None:
                parsed = self._to_canonical(item)
                # Filtre exact en Python : Restrict travaille en heure locale et
                # sur une fenetre volontairement large.
                if parsed is not None and parsed.end_utc > window.start and parsed.start_utc < window.end:
                    events.append(parsed)
                item = selection.GetNext()
        except Exception as exc:  # noqa: BLE001
            raise BackendError(self.key + ": parcours interrompu (" + str(exc) + ")") from exc

        return events

    def _to_canonical(self, item) -> CanonicalEvent | None:
        try:
            if item.Class != 26:  # olAppointment
                return None
            try:
                if int(item.MeetingStatus) in OL_CANCELLED_STATUS:
                    return None
            except Exception:  # noqa: BLE001 - statut illisible, on garde l'element
                pass
            all_day = bool(item.AllDayEvent)
            if all_day:
                first = _com_local_date(item.Start)
                last = _com_local_date(item.End)
                start_utc = datetime.combine(first, time.min, tzinfo=timezone.utc)
                end_utc = datetime.combine(max(last, first + timedelta(days=1)), time.min, tzinfo=timezone.utc)
                dates = (first, last)
            else:
                start_utc = _com_to_utc(item.StartUTC)
                end_utc = _com_to_utc(item.EndUTC)
                dates = None

            entry_id = item.EntryID
            try:
                recurring = bool(item.IsRecurring)
            except Exception:  # noqa: BLE001 - au moindre doute, on garde
                recurring = True  # l'horaire dans la cle : deux occurrences ne
                #                   doivent jamais se confondre.

            # Les marqueurs sont relus sur chaque element, sans pre-filtrage par
            # categorie : une copie dont la categorie aurait ete retiree a la
            # main redeviendrait sinon une source, et declencherait un echo.
            origin_raw = self._read_property(item, PROP_ORIGIN)
            stored_hash = self._read_property(item, PROP_HASH)

            return CanonicalEvent(
                cal=self.key,
                native_id=entry_id,
                series_id=entry_id,
                start_utc=start_utc,
                end_utc=end_utc,
                all_day=all_day,
                subject=item.Subject or "",
                location=item.Location or "",
                body=item.Body or "",
                busy=_BUSY_FROM_OL.get(int(item.BusyStatus), BUSY_BUSY),
                recurring=recurring,
                origin=OriginRef.parse(origin_raw or ""),
                stored_hash=stored_hash,
                all_day_dates=dates,
            )
        except Exception:  # noqa: BLE001 - element illisible, on l'ignore sans casser le run
            return None

    @staticmethod
    def _read_property(item, name: str) -> str | None:
        try:
            prop = item.UserProperties.Find(name)
        except Exception:  # noqa: BLE001
            return None
        if prop is None:
            return None
        try:
            value = prop.Value
        except Exception:  # noqa: BLE001
            return None
        return str(value) if value else None

    # -- ecriture ----------------------------------------------------------

    def _apply(self, appt, source: CanonicalEvent, origin: OriginRef) -> None:
        content = self.style.render(source)

        appt.Subject = content.subject
        appt.Location = content.location
        appt.Body = content.body
        appt.BusyStatus = _BUSY_TO_OL.get(content.busy, _BUSY_TO_OL[BUSY_BUSY])
        # Une copie n'est jamais une reunion : sans cela, Outlook enverrait des
        # invitations aux participants de l'evenement d'origine.
        appt.MeetingStatus = OL_NON_MEETING
        appt.ReminderSet = False
        appt.Categories = MIRROR_CATEGORY

        if source.all_day:
            if source.all_day_dates:
                first, last = source.all_day_dates
            else:
                first = source.start_utc.date()
                last = first + timedelta(days=1)
            if last <= first:
                last = first + timedelta(days=1)
            appt.AllDayEvent = True
            _set_wall_clock(appt, "Start", datetime.combine(first, time.min))
            _set_wall_clock(appt, "End", datetime.combine(last, time.min))
        else:
            appt.AllDayEvent = False
            # Start avant End : Outlook deplace End pour conserver la duree
            # quand on avance Start, et refuserait un End anterieur au Start.
            _set_utc(appt, "StartUTC", source.start_utc)
            _set_utc(appt, "EndUTC", source.end_utc)

        self._write_property(appt, PROP_ORIGIN, origin.key)
        self._write_property(appt, PROP_HASH, source.content_hash)
        appt.Save()

    @staticmethod
    def _write_property(appt, name: str, value: str) -> None:
        prop = appt.UserProperties.Find(name)
        if prop is None:
            # AddToFolderFields=True : rend le champ interrogeable par Restrict
            # et visible en colonne, ce qui aide au diagnostic manuel.
            prop = appt.UserProperties.Add(name, OL_TEXT, True)
        prop.Value = value

    def create(self, source: CanonicalEvent, origin: OriginRef) -> None:
        folder = self._connect()
        try:
            appt = folder.Items.Add(OL_APPOINTMENT_ITEM)
            self._apply(appt, source, origin)
        except Exception as exc:  # noqa: BLE001
            raise BackendError(self.key + ": creation impossible (" + str(exc) + ")") from exc

    def update(self, native_id: str, source: CanonicalEvent, origin: OriginRef) -> None:
        appt = self._item_by_id(native_id)
        if appt is None:
            # La copie a disparu entre la lecture et l'ecriture : le prochain run
            # la recreera proprement.
            return
        try:
            self._apply(appt, source, origin)
        except Exception as exc:  # noqa: BLE001
            raise BackendError(self.key + ": modification impossible (" + str(exc) + ")") from exc

    def delete(self, native_id: str) -> None:
        appt = self._item_by_id(native_id)
        if appt is None:
            return  # deja supprime : resultat voulu atteint
        try:
            appt.Delete()
        except Exception as exc:  # noqa: BLE001
            raise BackendError(self.key + ": suppression impossible (" + str(exc) + ")") from exc
