"""finance - DBS export -> scrub -> Claude -> expense tracker.

  finance run                 import new DBS exports, then categorise
  finance run --dry-run       save nothing, call nothing: writes the payload Claude would see
  finance categorise          categorise rows still marked NEEDS SUBCATEGORY
  finance categorise --retry  also re-ask rows Claude answered before (after editing [rules])
  finance status              health counters from the workbook
"""
from __future__ import annotations

import argparse
import collections
import datetime as dt
import glob
import json
import os
import re
import shutil
import sys
import tomllib
from pathlib import Path

from . import llm
from .dbs import read_export, row_key
from .scrub import Item, Pending, People, build_items, leak_check, payload
from .workbook import WorkbookOpen, excel_workbook

ROOT = Path(__file__).resolve().parent.parent
PRIVATE = ROOT / "private"
TEMPLATE = ROOT / "template" / "expense_tracker_template.xlsx"


def load_config() -> dict:
    return tomllib.loads((ROOT / "config.toml").read_text("utf-8"))


def load_state() -> dict:
    p = PRIVATE / "state.json"
    return json.loads(p.read_text("utf-8")) if p.exists() else {"imported": {}}


def save_state(state: dict):
    PRIVATE.mkdir(exist_ok=True)
    (PRIVATE / "state.json").write_text(json.dumps(state, indent=1), "utf-8")


def _date(s: str) -> dt.date | None:
    try:
        return dt.datetime.strptime(s.strip(), "%d-%b-%y").date()
    except ValueError:
        return None


# ---------------------------------------------------------------- import
def do_import(t, cfg, state, files: list[str] | None) -> int:
    pattern = os.path.expanduser(cfg["downloads_glob"])
    paths = [Path(f) for f in files] if files else [Path(p) for p in sorted(glob.glob(pattern))]
    existing = collections.Counter(row_key(r) for r in t.existing_keys())
    total = 0
    for p in paths:
        exp = read_export(p)
        if exp.sha256 in state["imported"]:
            continue
        new = []
        for r in exp.rows:
            k = row_key(r)
            if existing[k] > 0:
                existing[k] -= 1          # already in the workbook (overlapping export)
            else:
                new.append(r)
        dup = len(exp.rows) - len(new)
        if new:
            t.append(new)
            dates = [d for d in (_date(r[0]) for r in new) if d]
            t.log_import(min(dates).isoformat(), max(dates).isoformat(), len(new),
                         f"dbs-finance: {p.name}" + (f", {dup} duplicates skipped" if dup else ""))
        print(f"  {p.name}: {len(exp.rows)} rows, {len(new)} new, {dup} already present")
        state["imported"][exp.sha256] = {"file": p.name, "on": dt.date.today().isoformat(),
                                         "rows_added": len(new)}
        total += len(new)
    if not total:
        print("  no new transactions" + ("" if files else f" in files matching {pattern}"))
    return total


# ---------------------------------------------------------------- categorise
def build(t, cfg, people, retry=False) -> tuple[list[Item], list[dict]]:
    # Rows Claude already answered (Remarks start "AI") wait for you, unless --retry.
    pend = [Pending(**{k: v for k, v in p.items() if k != "remarks"}) for p in t.pending()
            if retry or not str(p["remarks"]).startswith("AI")]
    terms = t.private_terms()
    items = build_items(pend, people, cfg.get("scrub", {}).get("send_memos", False), terms)
    data = payload(items)
    problems = leak_check(data, people, terms)
    if problems:
        raise SystemExit("Refusing to send: scrubbed payload still contains\n  " + "\n  ".join(problems[:20]))
    return items, data


def _rule_pattern(it: Item, a: dict, descs: dict, taken: dict) -> str | None:
    """A Merchant_Map pattern is safe if it is in every description of this merchant,
    has no Excel wildcards, and matches no row already in a different subcategory."""
    for cand in (a.get("rule_text", "").strip(), it.pattern):
        if not cand or len(cand) < 4 or re.search(r"[*?~]", cand):
            continue
        c = " ".join(cand.split()).upper()
        if not all(c in descs[r][0] for r in it.rows):
            continue
        clash = {sub for d, sub in descs.values() if c in d and sub and sub != a["subcategory"]}
        if clash or taken.get(c, a["subcategory"]) != a["subcategory"]:
            continue
        return cand
    return None


def apply(t, items: list[Item], answers: list[dict]) -> collections.Counter:
    by_id = {i.id: i for i in items}
    today = dt.date.today().isoformat()
    descs = t.descriptions()
    taken: dict[str, str] = {}   # pattern (upper) -> subcategory, within this run
    rules, counts = [], collections.Counter()
    for a in answers:
        it, sub, conf, why = by_id[a["id"]], a["subcategory"], a["confidence"], a["reason"][:80]
        pattern = (_rule_pattern(it, a, descs, taken)
                   if conf == "high" and sub != "unknown" and not it.is_person else None)
        if sub == "unknown":
            for r in it.rows:
                t.set_override(r, None, f"AI: no idea - {why}")
            counts["unknown"] += len(it.rows)
        elif conf == "low":
            for r in it.rows:
                t.set_override(r, None, f"AI guess: {sub}? - {why}")
            counts["suggested"] += len(it.rows)
        elif pattern:
            key = " ".join(pattern.split()).upper()
            if key not in taken:          # one rule per pattern
                taken[key] = sub
                rules.append((pattern, sub, f"AI {today}: {why}"))
            counts["by_rule"] += len(it.rows)
        else:
            for r in it.rows:
                t.set_override(r, sub, f"AI ({conf}): {why}")
            counts["overridden"] += len(it.rows)
    t.add_rules(rules)
    counts["rules_added"] = len(rules)
    return counts


def do_categorise(t, cfg, retry=False, dry_run=False):
    people = People(PRIVATE / "people.json")
    items, data = build(t, cfg, people, retry)
    if not items:
        print("  nothing to categorise")
        return
    print(f"  {sum(i.count for i in items)} uncategorised rows -> {len(items)} distinct merchants after scrubbing")
    run_dir = ROOT / "runs" / dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    answers = llm.categorise(cfg.get("claude", {}), t.categories(), data,
                             cfg.get("rules", {}).get("notes", []), run_dir, dry_run)
    if dry_run:
        print(f"  dry run: payload and prompt written to {run_dir.relative_to(ROOT)}, nothing sent")
        return
    people.save()
    c = apply(t, items, answers)
    print(f"  applied: {c['by_rule']} rows via {c['rules_added']} new Merchant_Map rules, "
          f"{c['overridden']} rows set directly (flagged AI in Remarks), "
          f"{c['suggested']} low-confidence suggestions, {c['unknown']} unknown")


# ---------------------------------------------------------------- main
def main(argv=None):
    ap = argparse.ArgumentParser(prog="finance", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=["run", "categorise", "status"])
    ap.add_argument("files", nargs="*", help="DBS CSVs (default: downloads_glob in config.toml)")
    ap.add_argument("--dry-run", action="store_true", help="save nothing, call nothing")
    ap.add_argument("--retry", action="store_true", help="re-ask rows Claude already answered")
    args = ap.parse_intermixed_args(argv)   # flags may come before or after the files

    cfg = load_config()
    state = load_state()
    wb_path = (ROOT / cfg["workbook"]).resolve()
    if not wb_path.exists():
        wb_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(TEMPLATE, wb_path)
        print(f"Created {wb_path.relative_to(ROOT)} from the blank template")
    try:
        with excel_workbook(wb_path) as t:
            if args.command == "run":
                print("Importing")
                do_import(t, cfg, state, args.files)
                t.recalc()
            if args.command in ("run", "categorise"):
                print("Categorising with Claude" + (" (dry run)" if args.dry_run else ""))
                do_categorise(t, cfg, args.retry, args.dry_run)
                t.recalc()
            flags = t.check_flags()
            print("Health: " + (", ".join(f"{k} {v}" for k, v in sorted(flags.items()))
                                or "all rows categorised"))
            if args.command != "status":
                if args.dry_run:
                    print("Dry run: workbook not saved.")
                else:
                    t.save()
                    save_state(state)
                    print(f"Saved {wb_path.name}")
    except WorkbookOpen as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
