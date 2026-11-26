"""Ask Claude to pick a subcategory for each scrubbed merchant.

Runs Claude Code in headless mode (`claude -p`) on your Claude subscription login, so
no API key is needed. The reply is forced into a JSON schema.
"""
from __future__ import annotations

import json
import subprocess
import tempfile
from pathlib import Path

SYSTEM = (
    "You categorise personal bank transactions for a person living in Singapore. "
    "You only ever see a cleaned merchant label, never the raw bank record. "
    "Pick exactly one subcategory per merchant from the allowed list, or 'unknown'. "
    "Use your knowledge of Singapore merchants, chains, hawker centres and payment "
    "aggregators. Be honest about uncertainty."
)


def schema(allowed: list[str]) -> dict:
    return {
        "type": "object",
        "properties": {"items": {"type": "array", "items": {
            "type": "object",
            "properties": {
                "id": {"type": "string"},
                "subcategory": {"type": "string", "enum": allowed + ["unknown"]},
                "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
                "reason": {"type": "string"},
                "rule_text": {"type": "string"},
            },
            "required": ["id", "subcategory", "confidence", "reason", "rule_text"],
            "additionalProperties": False,
        }}},
        "required": ["items"],
        "additionalProperties": False,
    }


def build_prompt(categories: list[dict], items: list[dict], notes: list[str]) -> str:
    cat_lines = "\n".join(
        f"- {c['subcategory']}: {c['category']} / {c['type']} / {c['commitment']}" for c in categories)
    rules = "\n".join(f"- {n}" for n in notes)
    return f"""Categorise each merchant below.

Allowed subcategories (name: parent category / type / commitment):
{cat_lines}

Field meanings: statement_code POS = card payment, ADV = PayNow/bank transfer.
direction out = money left the account. A 3-letter code at the end of a card merchant
is the country (SGP Singapore, KOR Korea, MYS Malaysia ...).

Confidence:
- high: you know this merchant or the name is unambiguous. It will become a permanent rule.
- medium: probable. Applied to these rows only, flagged for review.
- low: a guess. Not applied, shown as a suggestion.

Owner's rules:
{rules}

rule_text: for high confidence card merchants, the shortest part of the merchant text
that names the brand for every outlet, copied exactly (e.g. "MCDONALD'S" from
"MCDONALD'S ORCHARD SI SGP"). Empty string when there is no clear brand, for
people, transfers, and anything not high confidence.

Return JSON matching the schema: one entry per id, reason under 12 words.

Merchants:
{json.dumps(items, indent=1, ensure_ascii=False)}
"""


def _validate(data: dict, ids: set[str], allowed: list[str]) -> list[dict]:
    out = []
    for a in data.get("items", []):
        if a.get("id") in ids and a.get("subcategory") in allowed + ["unknown"] \
                and a.get("confidence") in ("high", "medium", "low"):
            a.setdefault("rule_text", "")
            out.append(a)
    return out


def ask_claude(prompt: str, sch: dict, cfg: dict) -> dict:
    # Empty working directory, no tools, your user settings not loaded (so none of your
    # own hooks or plugins run) and no transcript saved.
    with tempfile.TemporaryDirectory() as empty:
        proc = subprocess.run(
            [cfg.get("command", "claude"), "-p", "--model", cfg.get("model", "opus"),
             "--tools", "", "--setting-sources", "project", "--no-session-persistence",
             "--output-format", "json", "--json-schema", json.dumps(sch),
             "--system-prompt", SYSTEM],
            input=prompt, capture_output=True, text=True, encoding="utf-8",
            cwd=empty, timeout=600)
    try:
        res = json.loads(proc.stdout)
    except json.JSONDecodeError:
        raise RuntimeError(f"claude CLI gave no JSON (exit {proc.returncode}): {proc.stderr[-500:]}")
    if res.get("is_error"):
        msg = res.get("result", "")
        if "authenticate" in msg.lower() or "log in" in msg.lower():
            msg += "\n  -> run once in a terminal:  claude auth login"
        raise RuntimeError(f"claude CLI error: {msg}")
    if isinstance(res.get("structured_output"), dict):
        return res["structured_output"]
    return json.loads(res["result"])


def categorise(cfg: dict, categories: list[dict], items: list[dict], notes: list[str],
               run_dir: Path, dry_run: bool = False) -> list[dict]:
    """Writes payload.json and prompt.md to run_dir, then asks Claude (unless dry_run)."""
    allowed = [c["subcategory"] for c in categories]
    sch = schema(allowed)
    run_dir.mkdir(parents=True, exist_ok=True)
    (run_dir / "payload.json").write_text(json.dumps(items, indent=1, ensure_ascii=False), "utf-8")
    (run_dir / "prompt.md").write_text(build_prompt(categories, items, notes), "utf-8")
    if dry_run:
        return []
    answers: list[dict] = []
    size = int(cfg.get("batch", 60))
    for start in range(0, len(items), size):
        chunk = items[start:start + size]
        data = ask_claude(build_prompt(categories, chunk, notes), sch, cfg)
        answers += _validate(data, {i["id"] for i in chunk}, allowed)
    (run_dir / "answers.json").write_text(json.dumps({"items": answers}, indent=1), "utf-8")
    missing = len(items) - len({a["id"] for a in answers})
    if missing:
        print(f"  note: no valid answer for {missing} merchant(s); they stay uncategorised")
    return answers
