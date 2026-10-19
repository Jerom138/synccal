"""Reconciliation des trois agendas.

Fonction pure : elle recoit les evenements deja lus et rend un plan d'actions.
Aucun appel COM, aucun appel reseau, donc entierement testable hors ligne.

Il n'existe deliberement aucun etat local : tout ce que le script sait de la
synchronisation precedente est relu depuis les marqueurs portes par les copies
elles-memes. Le script survit ainsi a un changement de poste, a une suppression
manuelle ou a une execution interrompue en plein milieu.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field

from core.model import CanonicalEvent, OriginRef, Window

CREATE = "create"
UPDATE = "update"
DELETE = "delete"


@dataclass(frozen=True)
class Action:
    kind: str
    target: str
    origin_key: str
    reason: str
    source: CanonicalEvent | None = None
    origin: OriginRef | None = None
    native_id: str | None = None
    victim: CanonicalEvent | None = None

    def describe(self) -> str:
        subject = ""
        if self.source is not None:
            subject = self.source.describe()
        elif self.victim is not None:
            subject = self.victim.describe()
        return f"[{self.kind:6}] {self.target:5} {subject}  ({self.reason})"


@dataclass
class Plan:
    actions: list[Action] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    suspicious: list[str] = field(default_factory=list)

    @property
    def deletions(self) -> int:
        return sum(1 for a in self.actions if a.kind == DELETE)

    def for_target(self, target: str) -> list[Action]:
        return [a for a in self.actions if a.target == target]

    def counts(self, target: str) -> dict[str, int]:
        result = {CREATE: 0, UPDATE: 0, DELETE: 0}
        for action in self.for_target(target):
            result[action.kind] += 1
        return result

    def is_empty(self) -> bool:
        return not self.actions


def build_plan(
    events_by_cal: dict[str, list[CanonicalEvent]],
    window: Window,
    writable: set[str] | None = None,
) -> Plan:
    """Compare l'etat souhaite a l'etat observe et en deduit les actions.

    `writable` liste les agendas qui acceptent l'ecriture. Un agenda en lecture
    seule ne genere aucune action : il diffuse vers les autres, mais ne recoit
    rien. Sans ce filtre, chaque run annoncerait des centaines de creations
    impossibles et le rapport deviendrait illisible.
    """
    plan = Plan()
    known_sources = set(events_by_cal)
    if writable is None:
        writable = set(events_by_cal)

    # Les evenements natifs sont les seules sources. Une copie n'est jamais
    # reprise comme source : c'est ce qui coupe la boucle d'echo entre les
    # trois agendas.
    natives: dict[str, list[CanonicalEvent]] = {
        cal: [ev for ev in events if not ev.is_mirror] for cal, events in events_by_cal.items()
    }

    mirrors: dict[str, dict[str, list[CanonicalEvent]]] = {}
    for cal, events in events_by_cal.items():
        grouped: dict[str, list[CanonicalEvent]] = defaultdict(list)
        for ev in events:
            if ev.is_mirror and ev.origin is not None:
                grouped[ev.origin.key].append(ev)
        # Tri stable : le survivant d'un doublon doit etre le meme d'un run a
        # l'autre, sans quoi le script oscillerait entre deux copies.
        for copies in grouped.values():
            copies.sort(key=lambda e: e.native_id)
        mirrors[cal] = grouped

    # Signature d'un agenda momentanement vide plutot que reellement vide :
    # il ne contient plus aucun evenement propre, alors que d'autres agendas
    # portent encore des copies qui le designent. C'est exactement l'etat d'un
    # profil Outlook en cours de premiere synchronisation, ou d'un OST reconstruit.
    # Appliquer le plan dans cet etat effacerait toutes ses copies.
    for cal in sorted(known_sources):
        if natives.get(cal):
            continue
        referencing = sum(
            len(copies)
            for other, grouped in mirrors.items()
            if other != cal
            for key, copies in grouped.items()
            if (OriginRef.parse(key) or OriginRef(cal="", series_id="", start_utc=None)).cal == cal
        )
        if referencing:
            plan.suspicious.append(
                f"{cal}: aucun evenement propre lu, alors que {referencing} copie(s) le designent ailleurs"
            )

    for target in sorted(events_by_cal):
        if target not in writable:
            manquantes = sum(
                1 for source_cal, events in natives.items() if source_cal != target for _ in events
            )
            if manquantes:
                plan.skipped.append(
                    f"{target}: agenda en lecture seule, {manquantes} copie(s) non placee(s)"
                )
            continue

        desired: dict[str, CanonicalEvent] = {}
        for source_cal, events in natives.items():
            if source_cal == target:
                continue
            for ev in events:
                desired[ev.self_ref.key] = ev

        present = mirrors[target]

        for key in sorted(desired):
            source = desired[key]
            copies = present.get(key, [])

            if not copies:
                plan.actions.append(
                    Action(
                        kind=CREATE,
                        target=target,
                        origin_key=key,
                        reason="copie absente",
                        source=source,
                        origin=source.self_ref,
                    )
                )
                continue

            keep, extras = copies[0], copies[1:]

            # Filet de securite : si un incident a laisse plusieurs copies pour
            # une meme source, les surnumeraires disparaissent ici.
            for extra in extras:
                plan.actions.append(
                    Action(
                        kind=DELETE,
                        target=target,
                        origin_key=key,
                        reason="doublon de la meme source",
                        native_id=extra.native_id,
                        victim=extra,
                    )
                )

            # Le hash decrit la source, pas la copie : une copie ecrite de
            # travers (bug d'ecriture, retouche manuelle) garde un hash
            # parfaitement valide et ne serait jamais reparee. Les horaires,
            # eux, doivent coincider quelle que soit la mise en forme des
            # copies : c'est l'invariant qui permet de detecter la derive.
            drifted = (
                keep.start_utc != source.start_utc
                or keep.end_utc != source.end_utc
                or keep.all_day != source.all_day
            )

            if drifted or keep.stored_hash != source.content_hash:
                plan.actions.append(
                    Action(
                        kind=UPDATE,
                        target=target,
                        origin_key=key,
                        reason="copie desynchronisee" if drifted else "source modifiee",
                        source=source,
                        origin=source.self_ref,
                        native_id=keep.native_id,
                        victim=keep,
                    )
                )

        for key in sorted(present):
            if key in desired:
                continue
            copies = present[key]
            ref = OriginRef.parse(key)

            if ref is None:
                plan.skipped.append(f"{target}: marqueur illisible, copie laissee intacte ({key!r})")
                continue

            # L'agenda source n'a pas ete lu ce run (desactive, ou en echec) :
            # on ne peut pas savoir si l'original existe encore. On ne touche rien.
            if ref.cal not in known_sources:
                plan.skipped.append(
                    f"{target}: {len(copies)} copie(s) de '{ref.cal}', agenda source non charge, laissees intactes"
                )
                continue

            # L'original est hors fenetre : son absence ne prouve pas sa
            # suppression. Protege les copies situees au bord de la fenetre.
            # Pour un evenement simple la cle ne porte pas d'horaire : on se
            # rabat sur celui de la copie, qui est le meme que celui de sa source.
            moment = ref.start_utc or copies[0].start_utc
            if not window.contains(moment):
                plan.skipped.append(f"{target}: copie hors fenetre laissee intacte ({moment:%Y-%m-%d})")
                continue

            for victim in copies:
                plan.actions.append(
                    Action(
                        kind=DELETE,
                        target=target,
                        origin_key=key,
                        reason="source supprimee",
                        native_id=victim.native_id,
                        victim=victim,
                    )
                )

    return plan
