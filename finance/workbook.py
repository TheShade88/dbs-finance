"""Read and write the expense-tracker workbook through Excel itself (COM).

openpyxl cannot be used to save this file: it drops the charts. Driving the installed
Excel keeps every chart, table, validation and formula, and recalculates for real, so
the script reads exactly what you would see. Windows with Excel installed is required.
"""
from __future__ import annotations

import re
from contextlib import contextmanager
from pathlib import Path

XL_UP = -4162
SHEET = "Transactions"


class WorkbookOpen(Exception):
    pass


@contextmanager
def excel_workbook(path: Path):
    import pythoncom
    import win32com.client

    path = path.resolve()
    lock = path.with_name("~$" + path.name)
    if lock.exists():
        raise WorkbookOpen(f"{path.name} is open in Excel. Close it and run again.")
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    wb = None
    try:
        wb = xl.Workbooks.Open(str(path))
        yield Tracker(xl, wb)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()
        pythoncom.CoUninitialize()


def _rows(value):
    if value is None:
        return []
    if not isinstance(value, tuple):
        return [[value]]
    return [list(r) for r in value]


class Tracker:
    def __init__(self, xl, wb):
        self.xl, self.wb = xl, wb
        self.tx = wb.Worksheets(SHEET)

    # ---------- basics ----------
    def last_row(self, ws=None, col=1) -> int:
        ws = ws or self.tx
        cell = ws.Cells(ws.Rows.Count, col).End(XL_UP)
        # End(xlUp) from below a table stops at the table's last row even when that
        # row is empty, so keep going up until a filled cell (or the header).
        while cell.Row > 1 and cell.Formula in ("", None):
            cell = cell.End(XL_UP)
        return cell.Row

    def recalc(self):
        self.xl.CalculateFull()

    def save(self):
        self.wb.Save()

    # ---------- transactions ----------
    def existing_keys(self) -> list[list]:
        last = self.last_row()
        if last < 2:
            return []
        return _rows(self.tx.Range(f"A2:L{last}").Value)

    def _ensure_capacity(self, need_last: int):
        lo = self.tx.ListObjects("tblTransactions")
        lo_last = lo.Range.Row + lo.Range.Rows.Count - 1
        if need_last <= lo_last:
            return
        new_last = need_last + 500
        lo.Resize(self.tx.Range(f"A1:AA{new_last}"))
        # Copy formulas down, column groups only: U and Z are your input columns.
        for cols in ("M:T", "V:Y", "AA:AA"):
            a, b = cols.split(":")
            self.tx.Range(f"{a}{lo_last}:{b}{new_last}").FillDown()

    def append(self, rows: list[list]) -> tuple[int, int]:
        start = self.last_row() + 1
        end = start + len(rows) - 1
        self._ensure_capacity(end)
        text = self.tx.Range(f"A{start}:J{end}")
        text.NumberFormat = "@"   # keep '27-Jul-26' as text: the date formulas parse it
        text.Value = [[("" if v is None else str(v)) for v in r[:10]] for r in rows]
        self.tx.Range(f"K{start}:L{end}").Value = [[("" if v is None else v) for v in r[10:12]]
                                                   for r in rows]
        return start, end

    def pending(self) -> list[dict]:
        """Rows the formulas could not categorise and you have not overridden."""
        last = self.last_row()
        vals = _rows(self.tx.Range(f"A2:AA{last}").Value)
        out = []
        for i, r in enumerate(vals, start=2):
            if r[26] == "NEEDS SUBCATEGORY" and not r[20]:
                out.append({"row": i, "code": r[2] or "", "desc": r[3] or "", "supp": r[4] or "",
                            "supp_desc": r[5] or "", "amount": float(r[15] or 0),
                            "fx_ccy": r[16] or "", "fx_amt": r[17], "remarks": r[25] or ""})
        return out

    def descriptions(self) -> dict[int, tuple[str, str]]:
        """row -> (description, final subcategory), for checking new rules."""
        last = self.last_row()
        vals = _rows(self.tx.Range(f"A2:V{last}").Value)
        return {i: (" ".join(str(r[3] or "").split()).upper(), r[21] or "")
                for i, r in enumerate(vals, start=2)}

    def private_terms(self) -> list[str]:
        """Distinctive words of whoever pays your salary (GR credits), so the employer's
        name never reaches the LLM. Found automatically; also picks up a new employer."""
        last = self.last_row()
        terms = set()
        for r in _rows(self.tx.Range(f"A2:P{last}").Value):
            if r[2] == "GR" and float(r[15] or 0) > 0 and r[4]:
                words = re.sub(r"^PAY\s+", "", " ".join(str(r[4]).split())).split()
                terms.update(w for w in words if len(w) >= 5 and w.upper() not in ("SALARY",))
        return sorted(terms)

    def check_flags(self) -> dict:
        last = self.last_row()
        flags = [r[0] for r in _rows(self.tx.Range(f"AA2:AA{last}").Value)]
        out = {}
        for f in flags:
            if f:
                out[f] = out.get(f, 0) + 1
        return out

    def set_override(self, row: int, subcategory: str | None, remark: str):
        if subcategory:
            self.tx.Range(f"U{row}").Value = subcategory
        z = self.tx.Range(f"Z{row}")
        if not z.Value or str(z.Value).startswith("AI"):
            z.Value = remark

    # ---------- reference sheets ----------
    def categories(self) -> list[dict]:
        ws = self.wb.Worksheets("Categories")
        out = []
        for r in _rows(ws.Range("A2:D200").Value):
            if not r[0]:
                break
            out.append({"subcategory": r[0], "category": r[1], "type": r[2], "commitment": r[3]})
        return out

    def add_rules(self, rules: list[tuple[str, str, str]]):
        """Insert Merchant_Map rules right below the existing ones (above the notes)."""
        if not rules:
            return
        ws = self.wb.Worksheets("Merchant_Map")
        r = 2
        while ws.Cells(r, 1).Value:
            r += 1
        n = len(rules)
        ws.Range(f"A{r}:C{r + n - 1}").EntireRow.Insert()
        ws.Range(f"A{r}:C{r + n - 1}").Value = [list(x) for x in rules]
        if r + n > 300:
            print("  warning: Merchant_Map passed row 300; widen the range in column T's formula")

    def log_import(self, period_from: str, period_to: str, added: int, notes: str):
        import datetime as dt
        ws = self.wb.Worksheets("Import_Log")
        r = 7
        while ws.Cells(r, 1).Value:
            r += 1
        ws.Range(f"A{r}:E{r}").EntireRow.Insert()
        cells = ws.Range(f"A{r}:E{r}")
        cells.NumberFormat = "@"
        cells.Value = [[dt.date.today().isoformat(), period_from, period_to, str(added), notes]]
