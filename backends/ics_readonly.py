"""Backend ICS en lecture seule, pour un agenda publie depuis Outlook Web.

C'est la porte de secours quand un agenda Microsoft 365 n'est joignable ni en
COM (compte non ajoutable au profil) ni en Graph (pas d'enregistrement
d'application). Elle ne permet que la lecture : les evenements de cet agenda
sont diffuses vers les deux autres, mais rien n'y est ecrit en retour.
"""

from __future__ import annotations

from dataclasses import replace
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path

from backends.base import BackendError, ReadOnlyBackend
from core.model import (
    BUSY_BUSY,
    BUSY_FREE,
    BUSY_OOF,
    BUSY_TENTATIVE,
    MIRROR_CATEGORY,
    CanonicalEvent,
    MirrorStyle,
    Window,
)

_TIMEOUT = 60

_BUSY_FROM_ICS = {
    "FREE": BUSY_FREE,
    "TENTATIVE": BUSY_TENTATIVE,
    "BUSY": BUSY_BUSY,
    "OOF": BUSY_OOF,
}


def _as_utc(value) -> tuple[datetime, bool, date | None]:
    """Normalise un DTSTART/DTEND iCalendar en instant UTC."""
    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.replace(tzinfo=timezone.utc, microsecond=0), False, None
        return value.astimezone(timezone.utc).replace(microsecond=0), False, None
    if isinstance(value, date):
        return datetime.combine(value, time.min, tzinfo=timezone.utc), True, value
    raise BackendError("ics: date illisible dans le flux")


def _text(component, field: str) -> str:
    raw = component.get(field)
    if raw is None:
        return ""
    return str(raw).strip()


class IcsReadOnlyBackend(ReadOnlyBackend):
    def __init__(self, key: str, label: str, settings: dict, style: MirrorStyle, root: Path) -> None:
        self.key = key
        self.label = label
        self.style = style
        self._url = (settings.get("url") or "").strip()
        # Un lien ICS publie est un secret de fait : il donne acces a l'agenda
        # sans authentification. On privilegie donc un fichier dans secrets/,
        # qui n'est pas versionne, plutot que config.yaml.
        url_file = (settings.get("url_file") or "").strip()
        if url_file:
            path = root / url_file
            if not path.exists():
                raise BackendError(
                    key + ": fichier de lien ICS introuvable (" + str(path) + ")"
                )
            self._url = path.read_text(encoding="utf-8").strip()

    def fetch(self, window: Window) -> list[CanonicalEvent]:
        if not self._url or self._url.startswith("REMPLACER"):
            raise BackendError(
                self.key + ": aucun lien ICS configure. Publier l'agenda depuis Outlook Web "
                "(Parametres > Calendrier > Calendriers partages > Publier un calendrier), "
                "niveau 'Tous les details', puis coller le lien .ics dans config.yaml."
            )

        try:
            import requests
        except ImportError as exc:
            raise BackendError(
                self.key + ": dependances manquantes, lancer pip install -r requirements.txt"
            ) from exc

        try:
            # requests honore HTTP_PROXY / HTTPS_PROXY : le proxy d'entreprise
            # est pris en compte sans configuration supplementaire.
            response = requests.get(self._url, timeout=_TIMEOUT)
            response.raise_for_status()
        except Exception as exc:  # noqa: BLE001
            raise BackendError(self.key + ": flux ICS injoignable (" + str(exc) + ")") from exc

        try:
            import icalendar
            import recurring_ical_events
        except ImportError as exc:
            raise BackendError(
                self.key + ": dependances manquantes, lancer pip install -r requirements.txt"
            ) from exc

        try:
            calendar = icalendar.Calendar.from_ical(response.content)
            # between() eclate les series recurrentes en occurrences, en tenant
            # compte des exceptions et des occurrences deplacees.
            occurrences = recurring_ical_events.of(calendar).between(window.start, window.end)
        except Exception as exc:  # noqa: BLE001
            raise BackendError(self.key + ": flux ICS illisible (" + str(exc) + ")") from exc

        events: list[CanonicalEvent] = []
        for component in occurrences:
            parsed = self._to_canonical(component)
            if parsed is not None:
                events.append(parsed)

        # Deuxieme passe : un UID qui revient plusieurs fois dans la fenetre est
        # forcement une serie, meme si la RRULE n'a pas survecu a l'expansion.
        # Mieux vaut marquer recurrent a tort (la copie sera recreee lors d'un
        # deplacement) que l'inverse, qui ferait s'ecraser les occurrences.
        seen: dict[str, int] = {}
        for event in events:
            seen[event.series_id] = seen.get(event.series_id, 0) + 1
        return [
            replace(event, recurring=True) if seen[event.series_id] > 1 else event
            for event in events
        ]

    def _to_canonical(self, component) -> CanonicalEvent | None:
        try:
            if str(component.get("STATUS", "")).upper() == "CANCELLED":
                return None

            # Garde-fou : si cet agenda a deja recu des copies par une autre
            # voie, elles ne doivent surtout pas etre rediffusees comme des
            # evenements d'origine. Le flux ICS ne transporte pas nos marqueurs,
            # mais il transporte la categorie.
            categories = component.get("CATEGORIES")
            if categories is not None and MIRROR_CATEGORY in str(categories.to_ical().decode("utf-8", "ignore")):
                return None

            start_utc, all_day, start_day = _as_utc(component.get("DTSTART").dt)

            end_prop = component.get("DTEND")
            if end_prop is not None:
                end_utc, _, end_day = _as_utc(end_prop.dt)
            else:
                duration = component.get("DURATION")
                span = duration.dt if duration is not None else timedelta(hours=1)
                end_utc, end_day = start_utc + span, None

            uid = _text(component, "UID") or (component.get("DTSTART").to_ical().decode())
            recurring = (
                component.get("RRULE") is not None
                or component.get("RECURRENCE-ID") is not None
                or component.get("RDATE") is not None
            )

            busy = BUSY_BUSY
            cdo = component.get("X-MICROSOFT-CDO-BUSYSTATUS")
            if cdo is not None:
                busy = _BUSY_FROM_ICS.get(str(cdo).upper(), BUSY_BUSY)
            elif str(component.get("TRANSP", "")).upper() == "TRANSPARENT":
                busy = BUSY_FREE

            dates = None
            if all_day and start_day:
                last = end_day or (start_day + timedelta(days=1))
                dates = (start_day, max(last, start_day + timedelta(days=1)))

            return CanonicalEvent(
                cal=self.key,
                native_id=uid + "@" + start_utc.strftime("%Y%m%dT%H%M%SZ"),
                series_id=uid,
                start_utc=start_utc,
                end_utc=end_utc,
                all_day=all_day,
                subject=_text(component, "SUMMARY"),
                location=_text(component, "LOCATION"),
                body=_text(component, "DESCRIPTION"),
                busy=busy,
                recurring=recurring,
                origin=None,  # rien ne nous appartient dans un agenda en lecture seule
                stored_hash=None,
                all_day_dates=dates,
            )
        except Exception:  # noqa: BLE001 - composant illisible, ignore sans casser le run
            return None
