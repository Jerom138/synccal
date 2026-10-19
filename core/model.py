"""Modele canonique partage par tous les backends.

Un evenement, quelle que soit sa provenance (COM, Google, ICS), est ramene a un
`CanonicalEvent`. Le moteur de reconciliation ne manipule que ce type : c'est ce
qui lui permet d'ignorer totalement la technologie d'acces sous-jacente.
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone

# Marqueurs poses sur les copies. Invisibles pour l'utilisateur, ils portent
# toute la memoire de la synchronisation : il n'y a volontairement aucune base
# de donnees locale.
PROP_ORIGIN = "SyncOrigin"
PROP_HASH = "SyncHash"
MIRROR_CATEGORY = "@Miroir"

# Valeurs normalisees de disponibilite.
BUSY_FREE = "free"
BUSY_TENTATIVE = "tentative"
BUSY_BUSY = "busy"
BUSY_OOF = "oof"
BUSY_ELSEWHERE = "elsewhere"

_KEY_TIME_FMT = "%Y%m%dT%H%M%SZ"
_TAG_RE = re.compile(r"<[^>]+>")
_WS_RE = re.compile(r"\s+")
_BODY_MAX = 4000
_SERIES_ID_MAX = 120
_NO_OCCURRENCE = "-"


def _compact(series_id: str) -> str:
    """Borne la longueur d'un identifiant de serie dans une cle d'origine.

    Les EntryID d'Outlook depassent facilement 200 caracteres, et une propriete
    etendue Google est limitee a 1024. Au-dela d'un seuil, l'identifiant est
    remplace par une empreinte : la cle reste comparable d'un run a l'autre, ce
    qui est tout ce que le moteur demande, et elle redevient bornee.
    """
    if len(series_id) <= _SERIES_ID_MAX:
        return series_id
    return "h:" + hashlib.sha1(series_id.encode("utf-8")).hexdigest()


class SyncError(RuntimeError):
    """Erreur fonctionnelle de synchronisation, remontee proprement au CLI."""


@dataclass(frozen=True)
class Window:
    """Fenetre temporelle de synchronisation, bornes en UTC."""

    start: datetime
    end: datetime

    def contains(self, moment: datetime) -> bool:
        return self.start <= moment < self.end

    @classmethod
    def build(cls, past_days: int, future_months: int, now: datetime | None = None) -> "Window":
        now = now or datetime.now(timezone.utc)
        # Approximation volontaire : un mois = 31 jours. L'horizon n'a pas
        # besoin d'etre exact, seulement genereux et stable d'un run a l'autre.
        return cls(
            start=(now - timedelta(days=past_days)).replace(microsecond=0),
            end=(now + timedelta(days=31 * future_months)).replace(microsecond=0),
        )

    def __str__(self) -> str:
        return f"{self.start:%Y-%m-%d} -> {self.end:%Y-%m-%d}"


@dataclass(frozen=True)
class OriginRef:
    """Reference vers l'evenement source d'une copie.

    Pour une occurrence de serie recurrente, l'horaire de debut fait partie de
    l'identite : via COM, toutes les occurrences partagent le meme EntryID (et
    le meme GlobalAppointmentID), et sans lui deux occurrences d'une meme
    reunion hebdomadaire s'ecraseraient mutuellement.

    Pour un evenement simple, au contraire, `start_utc` vaut None : son
    identifiant suffit a l'identifier, et le deplacer doit mettre la copie a
    jour plutot que la detruire pour en recreer une.
    """

    cal: str
    series_id: str
    start_utc: datetime | None

    @property
    def key(self) -> str:
        stamp = self.start_utc.strftime(_KEY_TIME_FMT) if self.start_utc else _NO_OCCURRENCE
        return f"{self.cal}|{_compact(self.series_id)}|{stamp}"

    @classmethod
    def parse(cls, raw: str) -> "OriginRef | None":
        """Relit une cle. Renvoie None si elle est illisible.

        Un marqueur corrompu ne doit jamais provoquer de suppression : l'appelant
        traite le None comme « origine inconnue, ne pas toucher ».
        """
        if not raw:
            return None
        parts = raw.split("|")
        if len(parts) < 3:
            return None
        cal, stamp = parts[0], parts[-1]
        series_id = "|".join(parts[1:-1])
        if stamp == _NO_OCCURRENCE:
            start = None
        else:
            try:
                start = datetime.strptime(stamp, _KEY_TIME_FMT).replace(tzinfo=timezone.utc)
            except ValueError:
                return None
        if not cal or not series_id:
            return None
        return cls(cal=cal, series_id=series_id, start_utc=start)


def normalize_body(raw: str | None) -> str:
    """Ramene un corps d'evenement a une forme comparable entre systemes.

    Outlook renvoie du HTML la ou Google renvoie du texte brut : sans cette
    normalisation, le hash differerait a chaque run et le script reecrirait
    indefiniment des evenements pourtant inchanges.
    """
    if not raw:
        return ""
    text = _TAG_RE.sub(" ", raw)
    text = text.replace("&nbsp;", " ").replace("&amp;", "&")
    text = text.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
    text = _WS_RE.sub(" ", text).strip()
    return text[:_BODY_MAX]


@dataclass(frozen=True)
class CanonicalEvent:
    """Un evenement normalise.

    `origin is None`  -> evenement natif, il sert de source.
    `origin` renseigne -> copie posee par le script, jamais reprise comme source.
    """

    cal: str
    native_id: str
    series_id: str
    start_utc: datetime
    end_utc: datetime
    all_day: bool
    subject: str = ""
    location: str = ""
    body: str = ""
    busy: str = BUSY_BUSY
    recurring: bool = False
    origin: OriginRef | None = None
    stored_hash: str | None = None
    all_day_dates: tuple[date, date] | None = field(default=None, compare=False)

    @property
    def is_mirror(self) -> bool:
        return self.origin is not None

    @property
    def self_ref(self) -> OriginRef:
        """Reference de cet evenement en tant que source."""
        return OriginRef(
            cal=self.cal,
            series_id=self.series_id,
            start_utc=self.start_utc if self.recurring else None,
        )

    @property
    def content_hash(self) -> str:
        """Empreinte du contenu, volontairement independante de l'agenda.

        Le meme evenement source produit la meme empreinte dans toutes ses
        copies : comparer l'empreinte stockee a celle de la source suffit pour
        savoir s'il faut reecrire.
        """
        payload = "\x1f".join(
            [
                (self.subject or "").strip(),
                self.start_utc.strftime(_KEY_TIME_FMT),
                self.end_utc.strftime(_KEY_TIME_FMT),
                "1" if self.all_day else "0",
                (self.location or "").strip(),
                normalize_body(self.body),
                self.busy,
            ]
        )
        return hashlib.sha1(payload.encode("utf-8")).hexdigest()

    def describe(self) -> str:
        """Ligne de rapport, en heure locale.

        Le moteur raisonne en UTC, mais l'utilisateur relit son agenda : afficher
        l'UTC ferait apparaitre 08:00 pour une reunion de 10:00.
        """
        if self.all_day:
            when = f"{self.start_utc:%Y-%m-%d} (journee)"
        else:
            when = f"{self.start_utc.astimezone():%Y-%m-%d %H:%M}"
        return f"{when} {self.subject or '(sans titre)'}"


@dataclass(frozen=True)
class MirrorContent:
    """Contenu effectif d'une copie, apres application des regles de confidentialite."""

    subject: str
    location: str
    body: str
    busy: str


@dataclass(frozen=True)
class MirrorStyle:
    """Regles de rendu des copies, communes a tous les backends.

    Centralisees ici pour qu'une copie soit rigoureusement identique quel que
    soit l'agenda de destination : c'est ce qui rend le hash de contenu
    comparable d'un agenda a l'autre.
    """

    subject_prefix: str = ""
    hide_subject: bool = False
    copy_location: bool = True
    copy_body: bool = True
    labels: dict = field(default_factory=dict)

    def render(self, source: CanonicalEvent) -> MirrorContent:
        label = self.labels.get(source.cal, source.cal)
        prefix = self.subject_prefix.format(label=label) if self.subject_prefix else ""

        if self.hide_subject:
            # Le titre reel ne sort pas de son agenda : le prefixe devient le
            # titre complet de la copie, et seul le creneau est partage.
            subject = prefix.strip() or "Occupe"
        else:
            subject = f"{prefix}{source.subject}".strip()

        return MirrorContent(
            subject=subject,
            location=source.location if self.copy_location else "",
            body=source.body if self.copy_body else "",
            busy=source.busy,
        )
