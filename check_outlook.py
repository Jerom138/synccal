"""Diagnostic Outlook classique - etape 0 de la mise en place.

Repond a une seule question : les agendas dont le script a besoin sont-ils
accessibles en COM sur ce poste ?

    python check_outlook.py

Le script ouvre Outlook classique s'il n'est pas deja lance. Si aucun profil
n'existe encore, Outlook affichera sa fenetre de creation de profil : c'est
normal, il faut la completer une fois, puis relancer ce diagnostic.
"""

from __future__ import annotations

import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

OL_FOLDER_CALENDAR = 9


def main() -> int:
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("pywin32 absent. Lancer : pip install -r requirements.txt")
        return 2

    try:
        pythoncom.CoInitialize()
    except Exception:  # noqa: BLE001
        pass

    print("Connexion a Outlook classique...")
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        namespace = app.GetNamespace("MAPI")
    except Exception as exc:  # noqa: BLE001
        print(f"ECHEC : Outlook classique injoignable ({exc})")
        print("Verifier que OUTLOOK.EXE demarre et qu'un profil existe.")
        return 1

    from backends.outlook_com import _restrict_format

    print(f"Format de date utilise pour les requetes : {_restrict_format()}")
    print("")

    stores = namespace.Stores
    print(f"{stores.Count} magasin(s) dans le profil :")
    print("")

    window_start = datetime.now(timezone.utc) - timedelta(days=30)
    window_end = datetime.now(timezone.utc) + timedelta(days=365)
    fmt = _restrict_format()
    usable = 0

    for index in range(1, stores.Count + 1):
        store = stores.Item(index)
        try:
            display = store.DisplayName
        except Exception as exc:  # noqa: BLE001
            print(f"  [{index}] magasin illisible ({exc})")
            continue

        try:
            folder = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
        except Exception:  # noqa: BLE001
            print(f"  [{index}] {display}")
            print("        pas de calendrier (archive ou magasin special)")
            continue

        try:
            items = folder.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True
            restriction = (
                "[Start] < '" + window_end.astimezone().strftime(fmt) + "' AND "
                "[End] > '" + window_start.astimezone().strftime(fmt) + "'"
            )
            selection = items.Restrict(restriction)
            count = 0
            item = selection.GetFirst()
            while item is not None and count < 5000:
                count += 1
                item = selection.GetNext()
            usable += 1
            print(f"  [{index}] {display}")
            print(f"        calendrier accessible, {count} evenement(s) dans la fenetre J-30 / J+365")
            print(f"        a reporter dans config.yaml  ->  store: \"{display}\"")
        except Exception as exc:  # noqa: BLE001
            print(f"  [{index}] {display}")
            print(f"        calendrier present mais illisible ({exc})")

    print("")
    if usable == 0:
        print("Aucun calendrier exploitable. Ajouter au moins un compte au profil Outlook.")
        return 1
    if usable == 1:
        print("Un seul calendrier accessible.")
        print("Pour synchroniser un second agenda en ecriture, ajouter son compte :")
        print("  Outlook classique > Fichier > Ajouter un compte")
        print("Si l'ajout est refuse, utiliser 'backend: ics_readonly' pour cet agenda.")
        return 0
    print(f"{usable} calendriers accessibles : la synchronisation complete est possible.")
    print("Renseigner le champ 'store' de chaque agenda dans config.yaml.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
