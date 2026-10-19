"""Autorisation Google Agenda - a lancer une seule fois.

    python -u authorize_google.py

Affiche l'URL d'autorisation, attend le retour, enregistre le jeton dans
secrets/token_google.json puis liste les calendriers du compte.

Pourquoi un script separe : l'ouverture automatique du navigateur echoue
silencieusement quand Python tourne dans un contexte non interactif. Ici l'URL
est toujours affichee, a coller manuellement si besoin.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from backends.google_cal import SCOPES  # noqa: E402

PORT = 8765
TOKEN = ROOT / "secrets" / "token_google.json"


def nom_attendu() -> str:
    """Nom du calendrier Google attendu, lu dans config.yaml.

    Evite de coder en dur un nom de calendrier dans le script : c'est la
    configuration qui fait foi, et elle seule.
    """
    try:
        import yaml

        with (ROOT / "config.yaml").open(encoding="utf-8") as handle:
            data = yaml.safe_load(handle) or {}
        for settings in (data.get("calendars") or {}).values():
            settings = settings or {}
            if settings.get("backend") == "google_cal":
                return (settings.get("calendar_name") or "").strip()
    except Exception:  # noqa: BLE001 - config absente ou illisible : on n'en fait pas un echec
        pass
    return ""


def main() -> int:
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError:
        print("Dependances manquantes. Lancer : pip install -r requirements.txt")
        return 2

    candidates = sorted((ROOT / "secrets").glob("client_secret*.json"))
    if not candidates:
        print("Aucun fichier client_secret*.json dans secrets/.")
        print("Telecharger le JSON 'Application de bureau' depuis Google Cloud Console.")
        return 2
    if len(candidates) > 1:
        print("Plusieurs fichiers client_secret*.json trouves, n'en garder qu'un :")
        for item in candidates:
            print("  -", item.name)
        return 2

    secret = candidates[0]
    print(f"Identifiants : {secret.name}")
    print("")

    flow = InstalledAppFlow.from_client_secrets_file(str(secret), SCOPES)

    print("=" * 78)
    print("Ouvrez l'URL ci-dessous dans votre navigateur (elle va aussi s'ouvrir seule).")
    print("Sur l'ecran 'Google n'a pas valide cette application' : Parametres avances,")
    print("puis 'Acceder a ... (non securise)'. C'est votre propre application.")
    print("=" * 78)
    print("")

    try:
        creds = flow.run_local_server(
            port=PORT,
            open_browser=True,
            authorization_prompt_message="{url}",
            success_message=(
                "Autorisation enregistree. Vous pouvez fermer cet onglet et "
                "revenir a Claude Code."
            ),
            timeout_seconds=600,
        )
    except Exception as exc:  # noqa: BLE001
        print("")
        print(f"Echec de l'autorisation : {exc}")
        print("")
        print("Si le navigateur affiche une erreur de connexion apres votre accord,")
        print("c'est que le proxy intercepte localhost. Configurer le navigateur pour")
        print("contourner le proxy sur 127.0.0.1, puis relancer ce script.")
        return 1

    TOKEN.parent.mkdir(parents=True, exist_ok=True)
    TOKEN.write_text(creds.to_json(), encoding="utf-8")
    print("")
    print(f"Jeton enregistre dans {TOKEN.relative_to(ROOT)}")
    print("")

    service = build("calendar", "v3", credentials=creds, cache_discovery=False)

    attendu = nom_attendu()
    if attendu:
        print(f"Calendrier attendu d'apres config.yaml : '{attendu}'")
    print("Calendriers accessibles sur ce compte :")
    token = None
    found = False
    while True:
        page = service.calendarList().list(pageToken=token, maxResults=250).execute(num_retries=5)
        for entry in page.get("items", []):
            summary = entry.get("summary", "")
            flag = ""
            if attendu and summary.strip().casefold() == attendu.casefold():
                flag = "   <-- celui attendu par config.yaml"
                found = True
            access = entry.get("accessRole", "?")
            print(f"  - {summary}  [{access}]{flag}")
        token = page.get("nextPageToken")
        if not token:
            break

    print("")
    if found:
        print(f"Le calendrier '{attendu}' existe : la configuration est bonne.")
        return 0
    if not attendu:
        print("Aucun 'calendar_name' lisible dans config.yaml : verifier le nom a la main.")
        return 0
    print(f"Aucun calendrier nomme exactement '{attendu}'.")
    print("Creer ce calendrier dans Google Agenda, ou corriger 'calendar_name' dans config.yaml.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
