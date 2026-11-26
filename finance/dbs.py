"""Read a DBS iBanking transaction-history CSV.

The export has a preamble (statement date, balances), then a header row starting with
"Transaction Date", then one row per transaction with 12 columns (A-L on the
Transactions sheet).
"""
from __future__ import annotations

import csv
import hashlib
import re
from dataclasses import dataclass
from pathlib import Path

HEADER = [
    "Transaction Date", "Value Date", "Statement Code", "Description",
    "Supplementary Code", "Supplementary Code Description", "Client Reference",
    "Additional Reference", "Status", "Currency", "Debit Amount", "Credit Amount",
]

# Full card number in any form DBS prints it (all digits, or already masked with X).
CARD_RE = re.compile(r"\b\d{4}-[\dX]{4}-[\dX]{4}-[\dX]{4}\b")
# Fixed-width token replacing every card number. Same length (19) as a real card
# number, so the workbook's character-offset formulas keep working, and it is the same
# for every card you will ever have, so a new card never breaks the spend-date parse.
CARD_TOKEN = "CARD-XXXX-XXXX-XXXX"
CARD_PREFIX = "CARD-"


def mask_cards(text: str) -> str:
    return CARD_RE.sub(CARD_TOKEN, text or "")


@dataclass
class Export:
    path: Path
    sha256: str
    statement_as_at: str
    rows: list[list]  # 12 fields each, card numbers already masked


def _num(s: str):
    s = (s or "").strip().replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return s


def read_export(path: Path) -> Export:
    raw = path.read_bytes()
    text = raw.decode("utf-8-sig", errors="replace")
    lines = list(csv.reader(text.splitlines()))
    try:
        h = next(i for i, r in enumerate(lines) if r and r[0].strip() == "Transaction Date")
    except StopIteration:
        raise ValueError(f"{path.name}: no 'Transaction Date' header row - not a DBS export?")
    if [c.strip() for c in lines[h][:12]] != HEADER:
        raise ValueError(f"{path.name}: header row changed, DBS may have altered the format:\n{lines[h]}")
    as_at = next((r[1] for r in lines[:h] if r and r[0].startswith("Statement as at")), "")
    rows = []
    for r in lines[h + 1:]:
        if not r or not r[0].strip():
            continue
        r = (r + [""] * 12)[:12]
        r = [mask_cards(c.strip()) if i < 10 else c for i, c in enumerate(r)]
        r[10], r[11] = _num(r[10]), _num(r[11])
        rows.append(r)
    return Export(path, hashlib.sha256(raw).hexdigest(), as_at, rows)


def row_key(r) -> tuple:
    """Identity of a transaction for duplicate detection. DBS has no transaction ID and
    Value Date is the export date, so use date + description + amounts."""
    def amt(v):
        return "" if v in (None, "") else f"{float(v):.2f}"
    return (str(r[0]).strip(), " ".join(str(r[3]).split()), amt(r[10]), amt(r[11]))
