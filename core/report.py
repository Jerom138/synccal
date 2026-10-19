"""Journalisation et rendu des rapports de synchronisation."""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

from core.engine import CREATE, DELETE, UPDATE, Plan

_LABELS = {CREATE: "cree", UPDATE: "modifie", DELETE: "supprime"}


def setup_logging(root: Path, verbose: bool) -> logging.Logger:
    """Console lisible, fichier complet.

    Le fichier garde toujours le detail : c'est la seule trace exploitable
    quand le script tourne depuis le planificateur de taches.
    """
    logs = root / "logs"
    logs.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("agendas")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    # La console Windows est en cp1252 : un emoji dans un titre d'evenement y
    # provoquerait une UnicodeEncodeError en pleine synchronisation. On degrade
    # les caracteres non representables plutot que d'interrompre le run.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001 - flux non reconfigurable (redirection, pipe)
        pass

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.DEBUG if verbose else logging.INFO)
    console.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(console)

    handler = logging.FileHandler(logs / f"sync-{datetime.now():%Y%m%d}.log", encoding="utf-8")
    handler.setLevel(logging.DEBUG)
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)-7s %(message)s"))
    logger.addHandler(handler)

    return logger


def render_inventory(logger: logging.Logger, inventory: dict[str, list]) -> None:
    logger.info("")
    logger.info("Inventaire")
    for cal, events in sorted(inventory.items()):
        natives = sum(1 for e in events if not e.is_mirror)
        mirrors = len(events) - natives
        logger.info(f"  {cal:5} {len(events):4} evenements  ({natives} propres, {mirrors} copies)")


def render_plan(logger: logging.Logger, plan: Plan, labels: dict[str, str]) -> None:
    logger.info("")
    if plan.is_empty():
        logger.info("Plan : rien a faire, les trois agendas sont alignes.")
    else:
        logger.info(f"Plan : {len(plan.actions)} action(s)")
        for cal in sorted({a.target for a in plan.actions}):
            counts = plan.counts(cal)
            logger.info(
                f"  {labels.get(cal, cal)} : "
                f"{counts[CREATE]} a creer, {counts[UPDATE]} a modifier, {counts[DELETE]} a supprimer"
            )
        for action in plan.actions:
            logger.debug("    " + action.describe())

    for note in plan.suspicious:
        logger.warning("  ATTENTION : " + note)

    for note in plan.skipped:
        logger.debug("  ignore : " + note)
    if plan.skipped:
        logger.info(f"  ({len(plan.skipped)} element(s) volontairement laisse(s) intact(s), detail en mode -v)")


def render_outcome(logger: logging.Logger, done: dict[str, int], errors: list[str]) -> None:
    logger.info("")
    if not any(done.values()):
        logger.info("Resultat : aucune ecriture.")
    else:
        parts = [f"{count} {_LABELS[kind]}" for kind, count in done.items() if count]
        logger.info("Resultat : " + ", ".join(parts) + ".")
    for message in errors:
        logger.error("  erreur : " + message)
    if errors:
        logger.error(f"{len(errors)} action(s) en echec. Relancer le script pour reprendre.")
