"""Test de bout en bout du cycle de synchronisation, sur agendas simules.

Les backends reels sont remplaces par des agendas en memoire qui se comportent
comme les vrais : ils stockent les marqueurs poses sur les copies et les
restituent a la lecture suivante. C'est ce qui permet de verifier la propriete
essentielle du script - relance a l'identique, il ne fait rien - sans toucher
a Outlook ni au reseau.
"""

from __future__ import annotations

import itertools
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.engine import CREATE, DELETE, UPDATE, build_plan
from core.model import CanonicalEvent, MirrorStyle, Window

NOW = datetime(2026, 10, 19, 12, 0, tzinfo=timezone.utc)
WIN = Window.build(past_days=30, future_months=12, now=NOW)
STYLE = MirrorStyle()

failures: list[str] = []
_ids = itertools.count(1)


def check(label: str, condition: bool, detail: str = "") -> None:
    if condition:
        print(f"  ok   {label}")
    else:
        failures.append(f"{label} {detail}".strip())
        print(f"  FAIL {label} {detail}")


class FakeCalendar:
    """Agenda en memoire respectant le contrat CalendarBackend."""

    def __init__(self, key: str, writable: bool = True) -> None:
        self.key = key
        self.label = key
        self.writable = writable
        self.store: dict[str, CanonicalEvent] = {}
        self.writes = 0

    def add_native(self, subject: str, offset_hours: int = 24) -> str:
        ident = f"{self.key}-n{next(_ids)}"
        start = NOW + timedelta(hours=offset_hours)
        self.store[ident] = CanonicalEvent(
            cal=self.key, native_id=ident, series_id=ident,
            start_utc=start, end_utc=start + timedelta(hours=1),
            all_day=False, subject=subject, location="Salle A", body="ordre du jour",
        )
        return ident

    def fetch(self, window):
        return list(self.store.values())

    def create(self, source, origin):
        self.writes += 1
        ident = f"{self.key}-m{next(_ids)}"
        self.store[ident] = self._mirror(ident, source, origin)

    def update(self, native_id, source, origin):
        self.writes += 1
        self.store[native_id] = self._mirror(native_id, source, origin)

    def delete(self, native_id):
        self.writes += 1
        self.store.pop(native_id, None)

    def _mirror(self, ident, source, origin) -> CanonicalEvent:
        content = STYLE.render(source)
        return CanonicalEvent(
            cal=self.key, native_id=ident, series_id=ident,
            start_utc=source.start_utc, end_utc=source.end_utc, all_day=source.all_day,
            subject=content.subject, location=content.location, body=content.body,
            busy=content.busy, origin=origin, stored_hash=source.content_hash,
        )

    def natives(self):
        return [e for e in self.store.values() if not e.is_mirror]

    def mirrors(self):
        return [e for e in self.store.values() if e.is_mirror]


def sync(cals: dict[str, FakeCalendar]) -> dict[str, int]:
    """Reproduit la boucle de sync.py : lecture, plan, execution."""
    inventory = {k: c.fetch(WIN) for k, c in cals.items()}
    plan = build_plan(inventory, WIN)
    done = {CREATE: 0, UPDATE: 0, DELETE: 0}
    for action in plan.actions:
        target = cals[action.target]
        if not target.writable:
            continue
        if action.kind == CREATE:
            target.create(action.source, action.origin)
        elif action.kind == UPDATE:
            target.update(action.native_id, action.source, action.origin)
        else:
            target.delete(action.native_id)
        done[action.kind] += 1
    return done


def fresh(readonly_pro_b: bool = False):
    return {
        "pro_a": FakeCalendar("pro_a"),
        "pro_b": FakeCalendar("pro_b", writable=not readonly_pro_b),
        "google": FakeCalendar("google"),
    }


print("\n1. Premiere synchro d'un maillage a 3")
cals = fresh()
cals["pro_a"].add_native("Reunion interne")
cals["pro_b"].add_native("Reunion client")
cals["google"].add_native("Perso")
done = sync(cals)
check("6 creations", done[CREATE] == 6, str(done))
check("chaque agenda a 1 propre + 2 copies",
      all(len(c.natives()) == 1 and len(c.mirrors()) == 2 for c in cals.values()),
      str({k: (len(c.natives()), len(c.mirrors())) for k, c in cals.items()}))

print("\n2. Idempotence sur 3 relances consecutives")
stable = True
for run in range(3):
    done = sync(cals)
    if any(done.values()):
        stable = False
        print(f"    run {run + 1} a produit {done}")
check("aucune action sur 3 relances", stable)
check("toujours 9 evenements au total", sum(len(c.store) for c in cals.values()) == 9)

print("\n3. Modification d'un original -> 2 mises a jour, aucune copie en plus")
pro_a_id = cals["pro_a"].natives()[0].native_id
original = cals["pro_a"].store[pro_a_id]
cals["pro_a"].store[pro_a_id] = CanonicalEvent(
    cal="pro_a", native_id=pro_a_id, series_id=pro_a_id,
    start_utc=original.start_utc + timedelta(hours=1),
    end_utc=original.end_utc + timedelta(hours=1),
    all_day=False, subject="Reunion interne (decale)", location="Salle B", body=original.body,
)
done = sync(cals)
check("2 mises a jour", done[UPDATE] == 2, str(done))
check("aucune creation", done[CREATE] == 0)
check("total inchange", sum(len(c.store) for c in cals.values()) == 9)
check("re-synchro stable", not any(sync(cals).values()))

print("\n4. Suppression d'un original -> ses 2 copies disparaissent")
del cals["pro_b"].store[cals["pro_b"].natives()[0].native_id]
done = sync(cals)
check("2 suppressions", done[DELETE] == 2, str(done))
check("6 evenements restants", sum(len(c.store) for c in cals.values()) == 6)
check("re-synchro stable", not any(sync(cals).values()))

print("\n5. Doublon injecte a la main -> rattrape au run suivant")
victim = cals["google"].mirrors()[0]
clone_id = "google-clone"
cals["google"].store[clone_id] = CanonicalEvent(
    cal="google", native_id=clone_id, series_id=clone_id,
    start_utc=victim.start_utc, end_utc=victim.end_utc, all_day=False,
    subject=victim.subject, origin=victim.origin, stored_hash=victim.stored_hash,
)
done = sync(cals)
check("1 suppression", done[DELETE] == 1, str(done))
check("plus aucun doublon",
      len({m.origin.key for m in cals["google"].mirrors()}) == len(cals["google"].mirrors()))
check("re-synchro stable", not any(sync(cals).values()))

print("\n6. Agenda en lecture seule : il diffuse mais ne recoit rien")
cals = fresh(readonly_pro_b=True)
cals["pro_a"].add_native("Comite")
cals["pro_b"].add_native("Reunion client")
sync(cals)
check("pro_b n a recu aucune ecriture", cals["pro_b"].writes == 0)
check("pro_b garde seulement son propre evenement", len(cals["pro_b"].store) == 1)
check("l evenement pro_b est bien diffuse ailleurs",
      len(cals["pro_a"].mirrors()) == 1 and len(cals["google"].mirrors()) == 2,
      f"pro_a={len(cals['pro_a'].mirrors())} google={len(cals['google'].mirrors())}")

print("\n7. Lecture seule : stabilite malgre le plan toujours non satisfiable")
before = {k: len(c.store) for k, c in cals.items()}
for _ in range(3):
    sync(cals)
check("aucune derive", {k: len(c.store) for k, c in cals.items()} == before,
      str({k: len(c.store) for k, c in cals.items()}))

print("\n8. Journee entiere : conservee et stable")
cals = fresh()
day = NOW.date() + timedelta(days=3)
cals["pro_a"].store["allday"] = CanonicalEvent(
    cal="pro_a", native_id="allday", series_id="allday",
    start_utc=datetime.combine(day, datetime.min.time(), tzinfo=timezone.utc),
    end_utc=datetime.combine(day + timedelta(days=1), datetime.min.time(), tzinfo=timezone.utc),
    all_day=True, subject="Conges", all_day_dates=(day, day + timedelta(days=1)),
)
done = sync(cals)
check("2 copies creees", done[CREATE] == 2)
check("copies marquees journee entiere", all(m.all_day for c in (cals["pro_b"], cals["google"]) for m in c.mirrors()))
check("re-synchro stable", not any(sync(cals).values()))

print()
if failures:
    print(f"{len(failures)} ECHEC(S)")
    for item in failures:
        print(f"  - {item}")
    sys.exit(1)
print("Cycle complet valide : convergence, idempotence, anti-doublon, lecture seule.")
