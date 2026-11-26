"""Parser and scrubber tests on an invented export. No Excel, no network."""
from pathlib import Path

import pytest

from finance.dbs import CARD_TOKEN, read_export, row_key
from finance.scrub import Pending, People, build_items, describe, leak_check, payload, redact

SAMPLE = Path(__file__).parent / "fixtures" / "transaction_history_sample.csv"


@pytest.fixture
def export():
    return read_export(SAMPLE)


def test_reads_rows_and_masks_cards(export):
    assert len(export.rows) == 11
    assert export.statement_as_at == "12-Aug-26"
    pos = export.rows[0]
    assert "4000-1234" not in pos[3] and CARD_TOKEN in pos[3]
    assert pos[6] == CARD_TOKEN
    assert pos[10] == 9.8 and pos[11] is None


def test_rejects_a_file_that_is_not_a_dbs_export(tmp_path):
    bad = tmp_path / "x.csv"
    bad.write_text("Date,Amount\n01-Aug-26,5\n")
    with pytest.raises(ValueError, match="not a DBS export"):
        read_export(bad)


def test_row_key_ignores_spacing_and_number_format(export):
    r = list(export.rows[0])
    r2 = list(r)
    r2[3] = r[3].replace(" ", "  ")
    r2[10] = "9.80"
    assert row_key(r) == row_key(r2)


def _pending(export):
    return [Pending(row=i, code=r[2], desc=r[3], supp=r[4], supp_desc=r[5],
                    amount=(r[11] or 0) - (r[10] or 0))
            for i, r in enumerate(export.rows, start=2)]


def test_card_merchant_label_drops_card_and_reference(export, tmp_path):
    label, pattern, is_person = describe(_pending(export)[0], People(tmp_path / "p.json"), False)
    assert label == "MCDONALD'S ORCHARD SI SGP"
    assert pattern == "MCDONALD'S ORCHARD"
    assert not is_person


def test_private_people_become_stable_tokens(export, tmp_path):
    people = People(tmp_path / "p.json")
    pend = _pending(export)
    out = describe(pend[5], people, False)
    assert out == ("PayNow to PERSON_01", None, True)
    assert describe(pend[7], people, False)[0] == "PayNow from PERSON_02"
    assert describe(pend[5], people, False)[0] == "PayNow to PERSON_01"   # same person, same token


def test_businesses_keep_their_name(export, tmp_path):
    label, pattern, is_person = describe(_pending(export)[6], People(tmp_path / "p.json"), False)
    assert label == "PayNow to EXAMPLE CAFE PTE LTD" and not is_person


def test_same_merchant_is_grouped_once(export, tmp_path):
    items = build_items(_pending(export), People(tmp_path / "p.json"))
    deca = [i for i in items if i.label.startswith("DECATHLON")]
    assert len(deca) == 1 and deca[0].count == 2 and deca[0].typical_amount_sgd == 28.95


def test_payload_passes_leak_check(export, tmp_path):
    people = People(tmp_path / "p.json")
    data = payload(build_items(_pending(export), people, private_terms=["WIDGETS"]))
    assert leak_check(data, people, ["WIDGETS"]) == []
    text = str(data)
    assert "TAN AH" not in text and "LIM MEI" not in text and "4000-1234" not in text


def test_leak_check_catches_what_slips_through(tmp_path):
    people = People(tmp_path / "p.json")
    people.token("TAN AH KOW")
    data = [{"merchant": "PAY TAN 91234567 ref 123456789"}]
    found = " ".join(leak_check(data, people))
    assert "long number" in found and "phone" in found


def test_employer_is_redacted():
    assert redact("RTF ACME WIDGETS PTE", ["WIDGETS"]) == "RTF ACME EMPLOYER"
