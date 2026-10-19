"""Backend Google Agenda (lecture/ecriture) via l'API Calendar v3."""

from __future__ import annotations

from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path

from backends.base import BackendError
from core.model import (
    BUSY_BUSY,
    BUSY_FREE,
    PROP_HASH,
    PROP_ORIGIN,
    CanonicalEvent,
    MirrorStyle,
    OriginRef,
    Window,
)

SCOPES = ["https://www.googleapis.com/auth/calendar"]
_PAGE_SIZE = 2500
# Reprise exponentielle integree au client Google : couvre 429, 5xx et les
# 403 rateLimitExceeded. Un premier run cree plusieurs centaines d'evenements
# d'affilee et depasse sinon la limite de debit par minute.
_RETRIES = 5


def _utc(moment: datetime) -> datetime:
    if moment.tzinfo is None:
        return moment.replace(tzinfo=timezone.utc)
    return moment.astimezone(timezone.utc)


def _parse_endpoint(node: dict) -> tuple[datetime, bool, date | None]:
    """Convertit un start/end Google en instant UTC.

    Une journee entiere est datee et non horodatee : on la rattache a minuit
    UTC, en conservant la date brute pour pouvoir la reecrire a l'identique.
    """
    if "date" in node:
        day = date.fromisoformat(node["date"])
        return datetime.combine(day, time.min, tzinfo=timezone.utc), True, day
    raw = node.get("dateTime")
    if not raw:
        raise BackendError("evenement Google sans date exploitable")
    return _utc(datetime.fromisoformat(raw)), False, None


class GoogleCalendarBackend:
    writable = True

    def __init__(self, key: str, label: str, settings: dict, style: MirrorStyle, root: Path) -> None:
        self.key = key
        self.label = label
        self.style = style
        self._calendar_name = (settings.get("calendar_name") or "").strip()
        if not self._calendar_name:
            raise BackendError(
                key + ": 'calendar_name' manquant dans config.yaml. Indiquer le nom exact "
                "du calendrier Google a synchroniser."
            )
        self._token = root / settings.get("token", "secrets/token_google.json")
        self._service = None
        self._calendar_id: str | None = None

    # -- connexion ---------------------------------------------------------

    def _credentials(self):
        """Charge l'autorisation Google, sans jamais ouvrir de navigateur.

        La synchronisation doit pouvoir tourner depuis le planificateur de
        taches : une invite interactive y resterait bloquee indefiniment. Toute
        (re)autorisation passe donc par authorize_google.py, lance a la main.
        """
        from google.auth.exceptions import RefreshError
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials

        if not self._token.exists():
            raise BackendError(
                self.key + ": aucune autorisation Google enregistree. "
                "Lancer une fois : python -u authorize_google.py"
            )

        try:
            creds = Credentials.from_authorized_user_file(str(self._token), SCOPES)
        except Exception as exc:  # noqa: BLE001
            raise BackendError(
                self.key + ": jeton illisible (" + str(exc) + "). "
                "Supprimer secrets/token_google.json puis relancer : python -u authorize_google.py"
            ) from exc

        if creds.valid:
            return creds

        if not creds.refresh_token:
            raise BackendError(
                self.key + ": jeton sans possibilite de renouvellement. "
                "Relancer : python -u authorize_google.py"
            )

        try:
            creds.refresh(Request())
        except RefreshError as exc:
            raise BackendError(
                self.key + ": autorisation Google expiree ou revoquee (" + str(exc) + "). "
                "Si l'ecran de consentement OAuth est en statut 'Test', Google invalide "
                "le jeton tous les 7 jours : relancer python -u authorize_google.py, ou "
                "publier l'application pour supprimer cette expiration."
            ) from exc

        self._token.write_text(creds.to_json(), encoding="utf-8")
        return creds

    def _connect(self):
        if self._service is not None:
            return self._service
        try:
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise BackendError(
                self.key + ": dependances manquantes, lancer pip install -r requirements.txt"
            ) from exc

        self._service = build(
            "calendar", "v3", credentials=self._credentials(), cache_discovery=False
        )
        self._calendar_id = self._resolve_calendar()
        return self._service

    def _resolve_calendar(self) -> str:
        wanted = self._calendar_name.strip().casefold()
        names: list[str] = []
        token = None
        while True:
            page = self._service.calendarList().list(pageToken=token, maxResults=250).execute(num_retries=_RETRIES)
            for entry in page.get("items", []):
                summary = entry.get("summary", "")
                names.append(summary)
                if summary.strip().casefold() == wanted:
                    return entry["id"]
            token = page.get("nextPageToken")
            if not token:
                break
        available = ", ".join(names) or "aucun"
        raise BackendError(
            self.key + ": calendrier '" + self._calendar_name + "' introuvable. "
            "Calendriers disponibles : " + available
        )

    # -- lecture -----------------------------------------------------------

    def fetch(self, window: Window) -> list[CanonicalEvent]:
        service = self._connect()
        events: list[CanonicalEvent] = []
        token = None
        try:
            while True:
                page = (
                    service.events()
                    .list(
                        calendarId=self._calendar_id,
                        timeMin=window.start.isoformat().replace("+00:00", "Z"),
                        timeMax=window.end.isoformat().replace("+00:00", "Z"),
                        singleEvents=True,  # recurrences deja eclatees en occurrences
                        showDeleted=False,
                        maxResults=_PAGE_SIZE,
                        pageToken=token,
                    )
                    .execute(num_retries=_RETRIES)
                )
                for item in page.get("items", []):
                    parsed = self._to_canonical(item)
                    if parsed is not None:
                        events.append(parsed)
                token = page.get("nextPageToken")
                if not token:
                    break
        except BackendError:
            raise
        except Exception as exc:  # noqa: BLE001 - toute panne de lecture est fatale
            raise BackendError(self.key + ": lecture impossible (" + str(exc) + ")") from exc
        return events

    def _to_canonical(self, item: dict) -> CanonicalEvent | None:
        if item.get("status") == "cancelled":
            return None
        try:
            start_utc, all_day, start_day = _parse_endpoint(item["start"])
            end_utc, _, end_day = _parse_endpoint(item["end"])
        except (BackendError, KeyError, ValueError):
            return None

        private = (item.get("extendedProperties") or {}).get("private") or {}
        origin = OriginRef.parse(private.get(PROP_ORIGIN, ""))

        dates = (start_day, end_day) if all_day and start_day and end_day else None

        return CanonicalEvent(
            cal=self.key,
            native_id=item["id"],
            series_id=item.get("recurringEventId") or item["id"],
            start_utc=start_utc,
            end_utc=end_utc,
            all_day=all_day,
            subject=item.get("summary") or "",
            location=item.get("location") or "",
            body=item.get("description") or "",
            busy=BUSY_FREE if item.get("transparency") == "transparent" else BUSY_BUSY,
            recurring=bool(item.get("recurringEventId")),
            origin=origin,
            stored_hash=private.get(PROP_HASH),
            all_day_dates=dates,
        )

    # -- ecriture ----------------------------------------------------------

    def _payload(self, source: CanonicalEvent, origin: OriginRef) -> dict:
        content = self.style.render(source)

        if source.all_day:
            if source.all_day_dates:
                first, last = source.all_day_dates
            else:
                first = source.start_utc.date()
                last = max(source.end_utc.date(), first + timedelta(days=1))
            start_node = {"date": first.isoformat()}
            end_node = {"date": last.isoformat()}
        else:
            start_node = {
                "dateTime": source.start_utc.isoformat().replace("+00:00", "Z"),
                "timeZone": "UTC",
            }
            end_node = {
                "dateTime": source.end_utc.isoformat().replace("+00:00", "Z"),
                "timeZone": "UTC",
            }

        return {
            "summary": content.subject,
            "location": content.location,
            "description": content.body,
            "start": start_node,
            "end": end_node,
            "transparency": "transparent" if content.busy == BUSY_FREE else "opaque",
            # Pas de rappel sur une copie : l'original porte deja le sien.
            "reminders": {"useDefault": False, "overrides": []},
            # Aucun participant : une copie ne doit jamais declencher d'invitation.
            "attendees": [],
            "extendedProperties": {
                "private": {PROP_ORIGIN: origin.key, PROP_HASH: source.content_hash}
            },
        }

    def create(self, source: CanonicalEvent, origin: OriginRef) -> None:
        service = self._connect()
        try:
            service.events().insert(
                calendarId=self._calendar_id,
                body=self._payload(source, origin),
                sendUpdates="none",
            ).execute(num_retries=_RETRIES)
        except Exception as exc:  # noqa: BLE001
            raise BackendError(self.key + ": creation impossible (" + str(exc) + ")") from exc

    def update(self, native_id: str, source: CanonicalEvent, origin: OriginRef) -> None:
        service = self._connect()
        try:
            service.events().update(
                calendarId=self._calendar_id,
                eventId=native_id,
                body=self._payload(source, origin),
                sendUpdates="none",
            ).execute(num_retries=_RETRIES)
        except Exception as exc:  # noqa: BLE001
            raise BackendError(self.key + ": modification impossible (" + str(exc) + ")") from exc

    def delete(self, native_id: str) -> None:
        service = self._connect()
        try:
            service.events().delete(
                calendarId=self._calendar_id, eventId=native_id, sendUpdates="none"
            ).execute(num_retries=_RETRIES)
        except Exception as exc:  # noqa: BLE001
            status = getattr(getattr(exc, "resp", None), "status", None)
            if status in (404, 410):
                return  # deja supprime cote Google : resultat voulu atteint
            raise BackendError(self.key + ": suppression impossible (" + str(exc) + ")") from exc
