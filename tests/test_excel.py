"""End-to-end through Excel on a copy of the blank template: import, recalculate, check
the formulas, build the payload. Nothing is saved and nothing is sent. Skipped when
Excel (pywin32) is not available. Takes about 20 s because Excel starts."""
import shutil
from pathlib import Path

import pytest

pytest.importorskip("win32com.client")

from finance import cli  # noqa: E402
from finance.scrub import People  # noqa: E402
from finance.workbook import excel_workbook  # noqa: E402

HERE = Path(__file__).parent
SAMPLE = HERE / "fixtures" / "transaction_history_sample.csv"


def test_import_categorise_payload(tmp_path):
    wb = tmp_path / "tracker.xlsx"
    shutil.copy(cli.TEMPLATE, wb)
    cfg = {"downloads_glob": str(tmp_path / "none_*.csv")}
    with excel_workbook(wb) as t:
        assert t.last_row() == 1
        assert cli.do_import(t, cfg, {"imported": {}}, [str(SAMPLE)]) == 11
        # same file again under a new hash: every row is already there
        again = tmp_path / "transaction_history_again.csv"
        again.write_bytes(SAMPLE.read_bytes() + b"\n")
        assert cli.do_import(t, cfg, {"imported": {}}, [str(again)]) == 0
        t.recalc()

        rows = t.tx.Range("A2:AA12").Value
        spend_dates = {r[3][:30]: r[13] for r in rows if r[2] == "POS"}
        assert all(d not in (None, "", "CHECK") for d in spend_dates.values()), spend_dates
        assert {r[14] for r in rows} == {"2026-08"}
        auto = {r[2] + " " + str(r[3] or "")[:12]: r[19] for r in rows}
        assert auto["POS BAT KOPITIAM"] == "food_hawker"        # Merchant_Map seed rule
        assert auto["GR PAY ACME WID"] == "income_salary"       # fixed rules
        assert auto["ATM CSH 12345678"] == "cash_withdrawal"
        assert auto["ATINT "] == "income_interest"

        assert "WIDGETS" in t.private_terms()
        items, data = cli.build(t, {}, People(tmp_path / "people.json"))
        labels = sorted(d["merchant"] for d in data)
        assert labels == ["DECATHLON SINGAPORE SI SGP", "EXAMPLE HOTEL SEOUL KOR",
                          "MCDONALD'S ORCHARD SI SGP", "PayNow from PERSON_02",
                          "PayNow to EXAMPLE CAFE PTE LTD", "PayNow to PERSON_01"]
        assert len(t.categories()) > 20
