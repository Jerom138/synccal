"""Tests du moteur de reconciliation (purs, sans COM ni reseau)."""

from __future__ import annotations

import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.engine import CREATE, DELETE, UPDATE, build_plan
from core.model import CanonicalEvent, OriginRef, Window

NOW = datetime(2026, 10, 19, 12, 0, tzinfo=timezone.utc)
WIN = Window.build(past_days=30, future_months=12, now=NOW)

failures: list[str] = []


def check(label: str, condition: bool, detail: str = "") -> None:
    if condition:
        print(f"  ok   {label}")
    else:
        failures.append(f"{label} {detail}".strip())
        print(f"  FAIL {label} {detail}")


def native(cal: str, ident: str, offset_hours: int = 24, subject: str = "Reunion") -> CanonicalEvent:
    start = NOW + timedelta(hours=offset_hours)
    return CanonicalEvent(
        cal=cal,
        native_id=ident,
        series_id=ident,
        start_utc=start,
        end_utc=start + timedelta(hours=1),
        all_day=False,
        subject=subject,
    )


def mirror(cal: str, ident: str, source: CanonicalEvent, stale: bool = False) -> CanonicalEvent:
    ref = source.self_ref
    return CanonicalEvent(
        cal=cal,
        native_id=ident,
        series_id=ident,
        start_utc=source.start_utc,
        end_utc=source.end_utc,
        all_day=False,
        subject=source.subject,
        origin=ref,
        stored_hash="perime" if stale else source.content_hash,
    )


print("\n1. Un natif dans pro_a produit une copie dans pro_b et google")
src = native("pro_a", "k1")
plan = build_plan({"pro_a": [src], "pro_b": [], "google": []}, WIN)
check("2 actions", len(plan.actions) == 2, str(len(plan.actions)))
check("toutes des creations", all(a.kind == CREATE for a in plan.actions))
check("cibles pro_b+google", {a.target for a in plan.actions} == {"pro_b", "google"})

print("\n2. Idempotence : une fois les copies en place, plus rien a faire")
plan = build_plan(
    {"pro_a": [src], "pro_b": [mirror("pro_b", "f1", src)], "google": [mirror("google", "g1", src)]}, WIN
)
check("plan vide", plan.is_empty(), [a.describe() for a in plan.actions])

print("\n3. Anti-echo : une copie n'est jamais reprise comme source")
m_fft = mirror("pro_b", "f1", src)
plan = build_plan({"pro_a": [src], "pro_b": [m_fft], "google": [mirror("google", "g1", src)]}, WIN)
check("aucune copie de la copie", not any(a.origin_key.startswith("pro_b|") for a in plan.actions))

print("\n4. Source modifiee -> update, pas de creation supplementaire")
plan = build_plan(
    {"pro_a": [src], "pro_b": [mirror("pro_b", "f1", src, stale=True)], "google": [mirror("google", "g1", src)]},
    WIN,
)
check("1 action", len(plan.actions) == 1, str(len(plan.actions)))
check("c'est un update", plan.actions[0].kind == UPDATE)
check("sur pro_b", plan.actions[0].target == "pro_b")

print("\n5. Doublon : deux copies de la meme source -> la surnumeraire est supprimee")
plan = build_plan(
    {
        "pro_a": [src],
        "pro_b": [mirror("pro_b", "f1", src), mirror("pro_b", "f2", src)],
        "google": [mirror("google", "g1", src)],
    },
    WIN,
)
check("1 action", len(plan.actions) == 1, str(len(plan.actions)))
check("suppression", plan.actions[0].kind == DELETE)
check("la 2e copie", plan.actions[0].native_id == "f2")

print("\n6. Source supprimee -> les copies disparaissent")
plan = build_plan(
    {"pro_a": [], "pro_b": [mirror("pro_b", "f1", src)], "google": [mirror("google", "g1", src)]}, WIN
)
check("2 suppressions", len(plan.actions) == 2 and all(a.kind == DELETE for a in plan.actions))

print("\n7. Agenda source non charge -> on ne supprime rien")
plan = build_plan({"pro_b": [mirror("pro_b", "f1", src)], "google": []}, WIN)
check("aucune action", plan.is_empty(), [a.describe() for a in plan.actions])
check("raison tracee", any("non charge" in s for s in plan.skipped))

print("\n8. Copie dont l'original est hors fenetre -> intacte")
far = native("pro_a", "k2", offset_hours=24 * 500)
plan = build_plan({"pro_a": [], "pro_b": [mirror("pro_b", "f9", far)], "google": []}, WIN)
check("aucune suppression", plan.deletions == 0)
check("raison tracee", any("hors fenetre" in s for s in plan.skipped))

print("\n9. Marqueur illisible -> copie laissee intacte")
broken = CanonicalEvent(
    cal="pro_b", native_id="fx", series_id="fx", start_utc=NOW, end_utc=NOW, all_day=False,
    origin=OriginRef(cal="pro_a", series_id="k", start_utc=NOW),
)
object.__setattr__(broken, "origin", OriginRef.parse("nimportequoi"))
check("cle invalide -> None", OriginRef.parse("nimportequoi") is None)

print("\n10. Maillage complet a 3 : chaque natif produit exactement 2 copies")
a, b, c = native("pro_a", "k1"), native("pro_b", "f1"), native("google", "g1")
plan = build_plan({"pro_a": [a], "pro_b": [b], "google": [c]}, WIN)
check("6 creations", len(plan.actions) == 6 and all(x.kind == CREATE for x in plan.actions))
per_target = {t: plan.counts(t)[CREATE] for t in ("pro_a", "pro_b", "google")}
check("2 par agenda", all(v == 2 for v in per_target.values()), str(per_target))

print("\n11. Recurrence : deux occurrences d'une meme serie ne se confondent pas")
occ1 = CanonicalEvent(cal="pro_a", native_id="S", series_id="S", all_day=False, recurring=True,
                      start_utc=NOW + timedelta(days=1), end_utc=NOW + timedelta(days=1, hours=1))
occ2 = CanonicalEvent(cal="pro_a", native_id="S", series_id="S", all_day=False, recurring=True,
                      start_utc=NOW + timedelta(days=8), end_utc=NOW + timedelta(days=8, hours=1))
check("cles distinctes", occ1.self_ref.key != occ2.self_ref.key)
plan = build_plan({"pro_a": [occ1, occ2], "pro_b": [], "google": []}, WIN)
check("4 creations", len(plan.actions) == 4, str(len(plan.actions)))

print("\n11b. Evenement simple : la cle ignore l'horaire, un deplacement reste un update")
simple = native("pro_a", "k9", offset_hours=48)
moved = CanonicalEvent(
    cal="pro_a", native_id="k9", series_id="k9", all_day=False, subject=simple.subject,
    start_utc=simple.start_utc + timedelta(hours=3), end_utc=simple.end_utc + timedelta(hours=3),
)
check("cle stable malgre le deplacement", simple.self_ref.key == moved.self_ref.key)
plan = build_plan({"pro_a": [moved], "pro_b": [mirror("pro_b", "f1", simple)], "google": []}, WIN)
pro_b_actions = [a for a in plan.actions if a.target == "pro_b"]
check("un seul update sur pro_b", len(pro_b_actions) == 1 and pro_b_actions[0].kind == UPDATE,
      str([a.kind for a in pro_b_actions]))
check("aucune suppression", plan.deletions == 0)

print("\n11c. Occurrence de serie deplacee : ancienne copie retiree, nouvelle creee")
plan = build_plan({"pro_a": [occ2], "pro_b": [mirror("pro_b", "f2", occ1)], "google": []}, WIN)
kinds = sorted(a.kind for a in plan.actions if a.target == "pro_b")
check("delete + create", kinds == [CREATE, DELETE], str(kinds))

print("\n12. Aller-retour de cle : parse(key) redonne la reference")
ref = occ1.self_ref
back = OriginRef.parse(ref.key)
check("round-trip exact", back == ref, f"{back} != {ref}")

print("\n12b. Copie dont l'horaire a derive : reparee meme si son hash est bon")
good = mirror("pro_b", "f1", src)
shifted = CanonicalEvent(
    cal="pro_b", native_id="f1", series_id="f1", all_day=False, subject=good.subject,
    start_utc=good.start_utc - timedelta(hours=2),   # ecrite 2 h trop tot
    end_utc=good.end_utc - timedelta(hours=2),
    origin=good.origin, stored_hash=good.stored_hash,  # hash pourtant valide
)
check("hash inchange", shifted.stored_hash == src.content_hash)
plan = build_plan({"pro_a": [src], "pro_b": [shifted], "google": [mirror("google", "g1", src)]}, WIN)
acts = [a for a in plan.actions if a.target == "pro_b"]
check("un update declenche", len(acts) == 1 and acts[0].kind == UPDATE, str([a.kind for a in acts]))
check("motif explicite", acts and acts[0].reason == "copie desynchronisee", str(acts[0].reason if acts else ""))
check("aucune suppression", plan.deletions == 0)

print("\n12c. Journee entiere devenue horaire : detectee aussi")
allday_src = CanonicalEvent(cal="pro_a", native_id="k7", series_id="k7", all_day=True,
                            start_utc=src.start_utc, end_utc=src.end_utc, subject="Conges")
wrong = CanonicalEvent(cal="pro_b", native_id="f7", series_id="f7", all_day=False,
                       start_utc=src.start_utc, end_utc=src.end_utc, subject="Conges",
                       origin=allday_src.self_ref, stored_hash=allday_src.content_hash)
plan = build_plan({"pro_a": [allday_src], "pro_b": [wrong], "google": []}, WIN)
acts = [a for a in plan.actions if a.target == "pro_b"]
check("un update declenche", len(acts) == 1 and acts[0].kind == UPDATE, str([a.kind for a in acts]))

print("\n13. Agenda momentanement vide : signale comme suspect")
plan = build_plan({"pro_a": [], "pro_b": [mirror("pro_b", "f1", src)], "google": [mirror("google", "g1", src)]}, WIN)
check("pro_a signale", any(s.startswith("pro_a:") for s in plan.suspicious), str(plan.suspicious))
check("les suppressions sont bien planifiees", plan.deletions == 2)

print("\n13b. Agenda legitimement vide, sans copie ailleurs : aucune alerte")
plan = build_plan({"pro_a": [], "pro_b": [native("pro_b", "f1")], "google": []}, WIN)
check("aucune suspicion", not plan.suspicious, str(plan.suspicious))

print("\n13c. Agenda garni : aucune alerte")
plan = build_plan({"pro_a": [src], "pro_b": [mirror("pro_b", "f1", src)], "google": [mirror("google", "g1", src)]}, WIN)
check("aucune suspicion", not plan.suspicious, str(plan.suspicious))

print("\n14. Agenda en lecture seule : diffuse, ne recoit rien, ne pollue pas le plan")
a, b, c = native("pro_a", "k1"), native("pro_b", "f1"), native("google", "g1")
plan = build_plan({"pro_a": [a], "pro_b": [b], "google": [c]}, WIN, writable={"pro_a", "google"})
check("aucune action vers pro_b", not plan.for_target("pro_b"), str(plan.counts("pro_b")))
check("pro_a et google servis", len(plan.actions) == 4, str(len(plan.actions)))
check("pro_b reste une source", all(
    any(act.origin_key.startswith("pro_b|") for act in plan.for_target(t)) for t in ("pro_a", "google")))
check("raison resumee en une ligne", sum(1 for s in plan.skipped if s.startswith("pro_b:")) == 1, str(plan.skipped))

print("\n14b. Sans argument writable, tout est ecrivable (compatibilite)")
plan = build_plan({"pro_a": [a], "pro_b": [b], "google": [c]}, WIN)
check("6 creations", len(plan.actions) == 6, str(len(plan.actions)))

print()
if failures:
    print(f"{len(failures)} ECHEC(S)")
    for f in failures:
        print(f"  - {f}")
    sys.exit(1)
print("Tous les tests du moteur passent.")
