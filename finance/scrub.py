"""Turn uncategorised transactions into the only thing an LLM ever sees.

What leaves the laptop, per distinct merchant: a cleaned merchant label, the DBS
statement code, money direction, a typical amount, how many times it occurred, and the
foreign currency if any. Never sent: card numbers, reference numbers, dates, balances,
salary, PayNow memos, or the names of private people (replaced by PERSON_nn tokens that
stay the same across months; the real names live only in private/people.json).

A final leak check refuses to build the payload if anything that looks like an account
number, card, phone, email or a known person's name survived the cleaning.
"""
from __future__ import annotations

import hashlib
import json
import re
import statistics
from dataclasses import dataclass, field
from pathlib import Path

# Words that mark a PayNow counterparty as a business. Anything without one is treated
# as a private person and pseudonymised. Erring this way only costs categorisation
# quality, never privacy.
BUSINESS_MARKERS = re.compile(
    r"\b(PTE|LTD|LLP|LLC|INC|CORP|CO|COMPANY|ENTERPRISES?|TRADING|HOLDINGS|GROUP|SERVICES?"
    r"|KIOSKS?|SHOP|STORE|MART|CAFE|COFFEE|RESTAURANT|FOOD|KITCHEN|BAKERY|CLINIC|DENTAL|MEDICAL"
    r"|PHARMACY|SALON|GYM|FITNESS|STUDIO|ACADEMY|SCHOOL|CENTRE|CENTER|HOTEL|TRAVEL|TOURS?"
    r"|PAY|SG|SINGAPORE|TOWN COUNCIL|MINISTRY|AUTHORITY|IRAS|HDB|SP SERVICES|TELCO|SINGTEL"
    r"|STARHUB|M1|CIRCLES|GIANT|NTUC|FAIRPRICE|SHENG SIONG|SOCIETY|CLUB|FOUNDATION|CHURCH|TEMPLE)\b"
    r"|[@&]|\.COM\b",
    re.I,
)

LEAK_PATTERNS = {
    "long number": re.compile(r"\d{6,}"),
    "card number": re.compile(r"\d{4}[- ]?\d{4}[- ]?\d{4}"),
    "email": re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"),
    "phone": re.compile(r"(?<!\d)[689]\d{7}(?!\d)|\+65"),
}


@dataclass
class Pending:
    row: int          # sheet row number (local only, never sent)
    code: str         # C Statement Code
    desc: str         # D Description
    supp: str         # E
    supp_desc: str    # F, e.g. "To: SOME MERCHANT"
    amount: float     # P signed amount, negative = money out
    fx_ccy: str = ""  # Q
    fx_amt: float | None = None  # R


@dataclass
class Item:
    id: str
    label: str
    code: str
    direction: str
    count: int
    typical_amount_sgd: float
    foreign_currency: str
    rows: list[int] = field(default_factory=list)
    pattern: str | None = None   # substring safe to add to Merchant_Map (local only)
    is_person: bool = False


class People:
    """Stable pseudonyms for private PayNow counterparties."""

    def __init__(self, path: Path):
        self.path = path
        self.map: dict[str, str] = json.loads(path.read_text("utf-8")) if path.exists() else {}

    def token(self, name: str) -> str:
        key = " ".join(name.upper().split())
        if key not in self.map:
            self.map[key] = f"PERSON_{len(self.map) + 1:02d}"
        return self.map[key]

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self.map, indent=1, ensure_ascii=False), "utf-8")


def _clean(text: str) -> str:
    t = re.sub(r"\bA-[A-Z0-9]{8,}\b", "", text)          # Grab booking ids
    t = re.sub(r"\b[A-Z0-9]*\d[A-Z0-9]{5,}\b", "", t)     # refs and long alphanumerics
    t = re.sub(r"\d{5,}", "", t)
    return " ".join(t.split()).strip(" ,-*#")


def _card_merchant(desc: str) -> tuple[str, str]:
    """POS descriptions: '<BAT> <merchant ... city CTY> <DDMMM> CARD-XXXX... <ref>'.
    Returns (merchant text as it appears in the description, 3-letter country)."""
    m = re.match(r"^[A-Z]{3}\s+(.*?)\s+\d{2}[A-Z]{3}\s+CARD-", desc)
    body = m.group(1) if m else re.sub(r"^[A-Z]{3}\s+", "", desc.split(" CARD-")[0])
    tail = re.search(r"\s([A-Z]{3})$", body)
    return body, tail.group(1) if tail else ""


def describe(p: Pending, people: People, send_memos: bool) -> tuple[str, str | None, bool]:
    """-> (label sent to the LLM, Merchant_Map pattern or None, is_person)"""
    desc = " ".join(p.desc.split())
    if p.code == "POS" or " CARD-" in desc:
        body, country = _card_merchant(desc)
        label = _clean(body)
        # Pattern for Merchant_Map: up to the first three words of the merchant, kept
        # contiguous so it is a true substring, with the city/country suffix dropped.
        words = body.split()
        while words and re.fullmatch(r"[A-Z]{2,3}|Si|SI", words[-1]):
            words.pop()
        pattern = " ".join(words[:3]) if words else None
        if pattern and (len(pattern) < 4 or re.search(r"\d{5,}", pattern)):
            pattern = None
        return label, pattern, False
    m = re.search(r"\b(To|From):\s*(.*?)(?:\s+OTHR\b(.*))?$", desc)
    if "PayNow" in desc and m:
        direction, party, memo = m.group(1), m.group(2).strip(), (m.group(3) or "").strip()
        is_person = not BUSINESS_MARKERS.search(party)
        name = people.token(party) if is_person else _clean(party)
        label = f"PayNow {'to' if direction == 'To' else 'from'} {name}"
        if send_memos and memo and memo.upper() != "PAYNOW TRANSFER" and not is_person:
            label += f" (memo: {_clean(memo)[:40]})"
        return label, (None if is_person else party), is_person
    if "NETS QR" in desc.upper():
        # NETS QR payees are always registered merchants
        party = re.split(r"TO:\s*", desc, flags=re.I)[-1].strip()
        return f"NETS QR payment to {_clean(party)}", party or None, False
    if desc.startswith("ICT "):
        # Transfer to another bank: the counterparty is usually a person, so send only
        # the purpose code DBS attaches (EDUC, RENT, OTHR ...).
        purpose = re.search(r"\b(EDUC|RENT|SALA|BEXP|SUPP|GDDS|TAXS|FAMI|LOAN|INVS|CHAR|MDCS"
                            r"|INSU|UTIL|IVPT|COMC|GIFT|OTHR)\b", desc)
        return ("Bank transfer to/from another bank"
                + (f", purpose code {purpose.group(1)}" if purpose else "")), None, True
    if desc.startswith("TRF"):
        return "Internal DBS transfer (TRF)", None, True
    words = [w for w in _clean(desc).split() if not re.search(r"\d.*\d.*\d|::", w)]
    return " ".join(words)[:60], None, False


def redact(label: str, terms: list[str]) -> str:
    for kw in terms:
        label = re.sub(rf"\b{re.escape(kw)}\b(\s+\S+){{0,2}}", "EMPLOYER", label, flags=re.I)
    return label


def build_items(pending: list[Pending], people: People, send_memos: bool = False,
                private_terms: list[str] = ()) -> list[Item]:
    groups: dict[tuple, Item] = {}
    amounts: dict[tuple, list[float]] = {}
    for p in pending:
        label, pattern, is_person = describe(p, people, send_memos)
        label = redact(label, private_terms)
        if "EMPLOYER" in label:
            pattern = None
        direction = "out" if p.amount < 0 else "in"
        key = (label.upper(), p.code, direction)
        if key not in groups:
            # letters only, so the leak check's number patterns never trip on an id
            iid = "m" + hashlib.sha1("|".join(key).encode()).hexdigest()[:8].translate(
                str.maketrans("0123456789", "ghijklmnop"))
            groups[key] = Item(iid, label, p.code, direction, 0, 0.0, p.fx_ccy or "",
                               pattern=pattern, is_person=is_person)
            amounts[key] = []
        it = groups[key]
        it.count += 1
        it.rows.append(p.row)
        amounts[key].append(abs(p.amount))
        # A pattern only survives if it appears in every description of the group, and
        # has no Excel SEARCH wildcards (* ? ~) that would make it match other rows.
        if it.pattern and (re.search(r"[*?~]", it.pattern)
                           or it.pattern.upper() not in " ".join(p.desc.split()).upper()):
            it.pattern = None
    for key, it in groups.items():
        it.typical_amount_sgd = round(statistics.median(amounts[key]), 2)
    return sorted(groups.values(), key=lambda i: -i.count)


def payload(items: list[Item]) -> list[dict]:
    return [{"id": i.id, "merchant": i.label, "statement_code": i.code, "direction": i.direction,
             "times_seen": i.count, "typical_amount_sgd": i.typical_amount_sgd,
             **({"foreign_currency": i.foreign_currency} if i.foreign_currency else {})}
            for i in items]


def leak_check(data: list[dict], people: People, private_terms: list[str] = ()) -> list[str]:
    text = json.dumps(data, ensure_ascii=False)
    problems = [f"{name}: {m.group(0)!r}" for name, rx in LEAK_PATTERNS.items()
                for m in rx.finditer(text)]
    problems += [f"private term: {t!r}" for t in private_terms
                 if re.search(rf"\b{re.escape(t)}\b", text, re.I)]
    for real in people.map:
        for part in real.split():
            if len(part) >= 4 and "X" not in part and re.search(rf"\b{re.escape(part)}\b", text, re.I):
                problems.append(f"person name fragment: {part!r}")
    return problems
