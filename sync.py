"""Synchronisation de trois agendas : deux Microsoft 365 et un Google.

Usage courant :
    python sync.py                 # simulation, aucune ecriture
    python sync.py --apply         # applique le plan
    python sync.py --apply -v      # avec le detail action par action

Le script est idempotent : le relancer sans rien changer ne produit aucune
action. C'est cette propriete qui garantit l'absence de doublon.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from backends.base import BackendError  # noqa: E402
from core.engine import CREATE, DELETE, UPDATE, Action, Plan, build_plan  # noqa: E402
from core.model import MirrorStyle, SyncError, Window  # noqa: E402
from core.report import render_inventory, render_outcome, render_plan, setup_logging  # noqa: E402

EXIT_OK = 0
EXIT_ERROR = 1
EXIT_CONFIG = 2
EXIT_GUARD = 3


def align_proxy_env() -> None:
    """Duplique les variables de proxy en minuscules.

    httplib2, utilise par le client Google, ne lit que http_proxy / https_proxy
    en minuscules, alors que Windows ne definit que les variantes en majuscules.
    Sans cet alignement, les appels Google partent hors proxy et echouent.
    """
    for name in ("HTTP_PROXY", "HTTPS_PROXY", "NO_PROXY"):
        value = os.environ.get(name)
        if value and not os.environ.get(name.lower()):
            os.environ[name.lower()] = value


def load_config(path: Path) -> dict:
    try:
        import yaml
    except ImportError as exc:
        raise SyncError("PyYAML absent, lancer pip install -r requirements.txt") from exc
    if not path.exists():
        raise SyncError(f"configuration introuvable : {path}")
    with path.open(encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}
    if not data.get("calendars"):
        raise SyncError(f"aucun agenda declare dans {path}")
    return data


def build_backends(config: dict, style: MirrorStyle):
    from backends.google_cal import GoogleCalendarBackend
    from backends.ics_readonly import IcsReadOnlyBackend
    from backends.outlook_com import OutlookComBackend

    registry = {
        "outlook_com": OutlookComBackend,
        "google_cal": GoogleCalendarBackend,
        "ics_readonly": IcsReadOnlyBackend,
    }

    backends = {}
    for key, settings in (config.get("calendars") or {}).items():
        settings = settings or {}
        if not settings.get("enabled", True):
            continue
        name = settings.get("backend")
        if name not in registry:
            raise SyncError(f"agenda '{key}' : backend inconnu '{name}'")
        backends[key] = registry[name](
            key=key,
            label=settings.get("label", key),
            settings=settings,
            style=style,
            root=ROOT,
        )

    if len(backends) < 2:
        raise SyncError("au moins deux agendas actifs sont necessaires pour synchroniser")
    return backends


def run(args: argparse.Namespace) -> int:
    align_proxy_env()

    config = load_config(Path(args.config) if args.config else ROOT / "config.yaml")
    logger = setup_logging(ROOT, args.verbose)

    window_cfg = config.get("window") or {}
    window = Window.build(
        past_days=int(window_cfg.get("past_days", 30)),
        future_months=int(window_cfg.get("future_months", 12)),
    )
    mirror_cfg = config.get("mirror") or {}
    style = MirrorStyle(
        subject_prefix=mirror_cfg.get("subject_prefix", "") or "",
        hide_subject=bool(mirror_cfg.get("hide_subject", False)),
        copy_location=bool(mirror_cfg.get("copy_location", True)),
        copy_body=bool(mirror_cfg.get("copy_body", True)),
        labels={
            key: (settings or {}).get("label", key)
            for key, settings in (config.get("calendars") or {}).items()
        },
    )
    max_deletions = int((config.get("safety") or {}).get("max_deletions", 20))

    backends = build_backends(config, style)
    labels = {key: backend.label for key, backend in backends.items()}

    mode = "APPLICATION" if args.apply else "SIMULATION"
    logger.info(f"Synchronisation agendas - {mode}")
    logger.info(f"Fenetre : {window}")
    logger.info("Agendas : " + ", ".join(
        f"{b.label} ({'lecture/ecriture' if b.writable else 'lecture seule'})"
        for b in backends.values()
    ))

    # Phase de lecture. Une seule defaillance condamne le run : un agenda
    # illisible ne doit jamais etre confondu avec un agenda vide, sous peine de
    # supprimer toutes ses copies dans les deux autres.
    inventory = {}
    for key, backend in backends.items():
        logger.debug(f"lecture de {backend.label}...")
        inventory[key] = backend.fetch(window)

    render_inventory(logger, inventory)

    if args.purge:
        # Marche arriere complete : on retire toutes les copies posees par le
        # script, sans toucher aux evenements d'origine. Indispensable pour
        # pouvoir essayer la synchronisation sans s'engager.
        plan = Plan()
        for key, backend in backends.items():
            if not backend.writable:
                continue
            for event in inventory[key]:
                if event.is_mirror and event.origin is not None:
                    plan.actions.append(
                        Action(
                            kind=DELETE,
                            target=key,
                            origin_key=event.origin.key,
                            reason="purge demandee",
                            native_id=event.native_id,
                            victim=event,
                        )
                    )
        logger.info("")
        logger.info("Mode purge : retrait de toutes les copies, les originaux sont conserves.")
    else:
        plan = build_plan(
            inventory, window, writable={k for k, b in backends.items() if b.writable}
        )

    render_plan(logger, plan, labels)

    readonly = [b.label for b in backends.values() if not b.writable]
    if readonly:
        logger.info("")
        logger.info(
            "En lecture seule, alimente les autres sans rien recevoir : " + ", ".join(sorted(readonly))
        )

    runnable = [a for a in plan.actions if backends[a.target].writable]

    if not args.apply:
        logger.info("")
        logger.info("Simulation terminee, aucune ecriture. Relancer avec --apply pour appliquer.")
        return EXIT_OK

    if plan.suspicious and not args.purge and not args.force:
        logger.error("")
        logger.error("Arret de securite : un agenda semble momentanement vide.")
        for note in plan.suspicious:
            logger.error("  " + note)
        logger.error(
            "Cause la plus frequente : profil Outlook en cours de premiere synchronisation, "
            "ou fichier OST reconstruit. Attendre la fin de la synchronisation puis relancer. "
            "Si cet agenda a bien ete vide volontairement, relancer avec --force."
        )
        return EXIT_GUARD

    deletions = sum(1 for a in runnable if a.kind == DELETE)
    if deletions > max_deletions and not args.purge and not args.force:
        logger.error("")
        logger.error(
            f"Arret de securite : {deletions} suppressions demandees (seuil {max_deletions}). "
            "Verifier le plan ci-dessus, puis relancer avec --force si ces suppressions sont voulues."
        )
        return EXIT_GUARD

    done = {CREATE: 0, UPDATE: 0, DELETE: 0}
    errors: list[str] = []
    for action in runnable:
        backend = backends[action.target]
        try:
            if action.kind == CREATE:
                backend.create(action.source, action.origin)
            elif action.kind == UPDATE:
                backend.update(action.native_id, action.source, action.origin)
            else:
                backend.delete(action.native_id)
            done[action.kind] += 1
            logger.debug("  fait : " + action.describe())
        except BackendError as exc:
            errors.append(f"{action.kind} sur {labels[action.target]} : {exc}")
            if len(errors) >= 10:
                logger.error("Trop d'echecs consecutifs, arret du run.")
                break

    render_outcome(logger, done, errors)
    return EXIT_ERROR if errors else EXIT_OK


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Synchronise les agendas declares dans config.yaml, sans doublon."
    )
    parser.add_argument("--apply", action="store_true", help="applique le plan (defaut : simulation)")
    parser.add_argument("--force", action="store_true", help="passe outre le garde-fou de suppressions")
    parser.add_argument(
        "--purge",
        action="store_true",
        help="retire toutes les copies posees par le script, sans toucher aux originaux",
    )
    parser.add_argument("--config", help="chemin d'un fichier de configuration alternatif")
    parser.add_argument("-v", "--verbose", action="store_true", help="detaille chaque action")
    args = parser.parse_args()

    try:
        return run(args)
    except SyncError as exc:
        print(f"Configuration : {exc}", file=sys.stderr)
        return EXIT_CONFIG
    except BackendError as exc:
        print(f"Acces agenda : {exc}", file=sys.stderr)
        print("Aucune modification n'a ete faite.", file=sys.stderr)
        return EXIT_ERROR
    except KeyboardInterrupt:
        print("Interrompu.", file=sys.stderr)
        return EXIT_ERROR


if __name__ == "__main__":
    raise SystemExit(main())
