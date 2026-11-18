"""Builder for the MDK CBE weekly planning workbook (approved template v3).

This is a faithful port of ``build()`` / ``read_master()`` from ``planning_template_reference.py``
(the approved generator). Formulas, styles, sheet order, validations, conditional formats,
widths, filters and freeze panes are identical. The only deliberate differences:

1. Text that comes from the master (work centre and item names, and the Consolidated copy of the
   master) is written as a literal string even when it starts with ``= + - @``, so uploaded text
   can never become a formula (formula-injection protection). A cell that is a real formula in
   the master is still copied as a formula, exactly as the reference does.
2. It works on in-memory file objects and returns a typed :class:`BuildInfo`.
3. Header detection tolerates case and extra whitespace (see :mod:`planner.headers`).

Golden tests (``tests/test_golden.py``) prove the output equals the reference generator's.
"""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from typing import IO, Any

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L  # noqa: N812 - name kept from the reference
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation

from planner.errors import MasterFormatError
from planner.headers import HEADER_SEARCH_ROWS, find_header_row

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
#: Sheet names (case-insensitive) that the builder deletes and recreates.
RESERVED_SHEETS = [d.lower() for d in DAYS] + ["consolidated", "_lists"]
FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
WC_FILL = PatternFill("solid", fgColor="D9E2F3")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

MasterRow = tuple[int, Any, Any, Any]


@dataclass(frozen=True)
class BuildInfo:
    """Layout facts about a generated workbook (used for the summary and tests).

    Attributes:
        wcs: Distinct work centres in first-appearance order.
        items: Work centre -> its distinct items, in master order.
        first: First data row on the day sheets.
        lastd: Last data row on the day sheets.
        hdr: Header row number (same on master, day sheets and Consolidated).
        last: Last non-blank master row.
        nrows: Number of planning rows per day sheet.
        rows: Non-blank master rows as ``(row, item, work centre, qty)``.
        master_title: Title of the master sheet.

    """

    wcs: list[Any]
    items: dict[Any, list[Any]]
    first: int
    lastd: int
    hdr: int
    last: int
    nrows: int
    rows: list[MasterRow]
    master_title: str


def _put(ws: Any, row: int, col: int, value: Any) -> Any:
    """Write a master-derived value; text is always stored literally, never as a formula."""
    cell = ws.cell(row, col)
    cell.value = value
    if isinstance(value, str) and cell.data_type == "f":
        cell.data_type = "s"
    return cell


def read_master(wb: Any) -> tuple[Any, int, int, list[MasterRow]]:
    """Locate the header row on the first sheet and return its non-blank data rows."""
    ws = wb.worksheets[0]
    top = [tuple(ws.cell(r, c).value for c in range(1, 4)) for r in range(1, HEADER_SEARCH_ROWS + 1)]
    hdr = find_header_row(top)
    if hdr is None:
        raise MasterFormatError("Header row (Item Name | Work Center | Week Plan Qty) not found")
    rows: list[MasterRow] = []
    last = hdr
    for r in range(hdr + 1, ws.max_row + 1):
        item, wc, qty = (ws.cell(r, c).value for c in (1, 2, 3))
        if item is None and wc is None and qty is None:
            continue
        last = r
        rows.append((r, item, wc, qty))
    return ws, hdr, last, rows


def build(src: IO[bytes] | str, dst: IO[bytes] | str) -> BuildInfo:
    """Build the planning workbook from the master at ``src`` and save it to ``dst``.

    Raises:
        MasterFormatError: no header row, or no row with both an item and a work centre.

    """
    wb = load_workbook(src)
    # replace pre-existing day tabs if any
    for name in list(wb.sheetnames):
        if name.lower() in RESERVED_SHEETS:
            del wb[name]
    if not wb.worksheets:
        raise MasterFormatError("The workbook has no master sheet.")
    master, hdr, last, rows = read_master(wb)
    mname = master.title
    mq = "'%s'" % mname.replace("'", "''")
    title = master["A1"].value or ""

    # distinct work centres (first appearance) and their items, master order, as-is
    wcs: list[Any] = []
    items: dict[Any, list[Any]] = {}
    for _, item, wc, _qty in rows:
        if wc is None or item is None:
            continue
        if wc not in items:
            wcs.append(wc)
            items[wc] = []
        if item not in items[wc]:
            items[wc].append(item)
    if not wcs:
        raise MasterFormatError(
            "No rows with both an Item Name and a Work Center were found below the header."
        )

    # hidden lists sheet
    lists = wb.create_sheet("_Lists")
    for j, wc in enumerate(wcs, start=1):
        _put(lists, 1, j, wc)
        for i, it in enumerate(items[wc], start=2):
            _put(lists, i, j, it)
    lists.sheet_state = "hidden"
    maxitems = max(len(v) for v in items.values())
    hdr_rng = "_Lists!$A$1:$%s$1" % L(len(wcs))

    first = hdr + 1  # 4
    nrows = sum(len(items[w]) for w in wcs)
    lastd = first + nrows - 1
    M_A = "%s!$A$%d:$A$%d" % (mq, hdr + 1, last)  # noqa: N806 - names kept from the reference
    M_B = "%s!$B$%d:$B$%d" % (mq, hdr + 1, last)  # noqa: N806
    M_C = "%s!$C$%d:$C$%d" % (mq, hdr + 1, last)  # noqa: N806

    for di, day in enumerate(DAYS):
        ws = wb.create_sheet(day)
        ws["A1"] = "%s  |  %s" % (day, title)
        ws["A1"].font = Font(name=FONT, bold=True, size=13)
        ws["A2"] = (
            "Pick an Item from the dropdown, then type Day Plan Qty (and Priority). "
            "Red Day Plan Qty = more than Available. Orange Item = same item picked twice today. "
            "Yellow cells are for you to fill; other columns are formulas. "
            "Each work centre has one row per item it can plan; pick a different item on each row. "
            ""
        )
        ws["A2"].font = Font(name=FONT, italic=True, size=9, color="595959")
        ws.merge_cells("A2:G2")
        ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[2].height = 38
        heads = [
            "Work centre",
            "Item",
            "Weekly plan",
            "Available",
            "Day Plan Qty",
            "Balance to be planned",
            "Priority",
        ]
        for c, h in enumerate(heads, start=1):
            cell = ws.cell(hdr, c, h)
            cell.font = Font(name=FONT, bold=True, color="FFFFFF")
            cell.fill = HDR_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER
        ws.row_dimensions[hdr].height = 30

        prev = DAYS[:di]
        r = first
        for wc in wcs:
            main = r
            for k in range(len(items[wc])):
                if k == 0:
                    _put(ws, r, 1, wc)
                else:
                    ws.cell(r, 1, "=$A$%d" % main)
                    ws.cell(r, 1).number_format = ";;;"
                # weekly plan
                ws.cell(
                    r,
                    3,
                    '=IF($B{r}="","",SUMPRODUCT(({a}=$B{r})*({b}=$A{r})*{c}))'.format(
                        r=r, a=M_A, b=M_B, c=M_C
                    ),
                )
                # available
                if not prev:
                    avail = '=IF($B{r}="","",$C{r})'.format(r=r)
                else:
                    sub = "".join(
                        "-SUMPRODUCT(({d}!$A${f}:$A${l}=$A{r})*({d}!$B${f}:$B${l}=$B{r})*{d}!$E${f}:$E${l})".format(
                            d=p, f=first, l=lastd, r=r
                        )
                        for p in prev
                    )
                    avail = '=IF($B{r}="","",$C{r}{sub})'.format(r=r, sub=sub)
                ws.cell(r, 4, avail)
                ws.cell(r, 6, '=IF($B{r}="","",$D{r}-N($E{r}))'.format(r=r))
                for c in range(1, 8):
                    cell = ws.cell(r, c)
                    cell.font = Font(name=FONT, size=10, bold=(c == 1 and k == 0))
                    cell.border = BORDER
                    if c in (2, 5, 7):
                        cell.fill = INPUT_FILL
                    if c == 1 and k == 0:
                        cell.fill = WC_FILL
                    if c in (3, 4, 5, 6):
                        cell.number_format = "#,##0"
                ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="center")
                r += 1

        # dropdown for items depending on work centre
        dv = DataValidation(
            type="list",
            formula1=(
                "OFFSET(_Lists!$A$2,0,MATCH($A{f},{h},0)-1,"
                "COUNTA(OFFSET(_Lists!$A$2,0,MATCH($A{f},{h},0)-1,{m},1)),1)"
            ).format(f=first, h=hdr_rng, m=maxitems),
            allow_blank=True,
            showErrorMessage=True,
            showInputMessage=True,
            errorTitle="Invalid item",
            error="Pick an item from the list for this work centre.",
            promptTitle="Item",
            prompt="Choose an item for this work centre.",
        )
        dv.add("B%d:B%d" % (first, lastd))
        ws.add_data_validation(dv)
        dq = DataValidation(
            type="decimal",
            operator="greaterThanOrEqual",
            formula1="0",
            allow_blank=True,
            showErrorMessage=True,
            errorTitle="Invalid quantity",
            error="Enter a number, 0 or more.",
        )
        dq.add("E%d:E%d" % (first, lastd))
        ws.add_data_validation(dq)

        ws.conditional_formatting.add(
            "E%d:E%d" % (first, lastd),
            FormulaRule(
                formula=['AND($B%d<>"",N($E%d)>N($D%d))' % (first, first, first)],
                fill=PatternFill("solid", bgColor="FF0000", fgColor="FF0000"),
                font=Font(color="FFFFFF", bold=True),
            ),
        )
        ws.conditional_formatting.add(
            "B%d:B%d" % (first, lastd),
            FormulaRule(
                formula=[
                    'AND($B{f}<>"",SUMPRODUCT(($A${f}:$A${l}=$A{f})*($B${f}:$B${l}=$B{f}))>1)'.format(
                        f=first, l=lastd
                    )
                ],
                fill=PatternFill("solid", bgColor="F4B183", fgColor="F4B183"),
            ),
        )

        for col, w in zip("ABCDEFG", (24, 58, 13, 13, 14, 16, 12), strict=True):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = ws.cell(first, 1)
        ws.auto_filter.ref = "A%d:G%d" % (hdr, lastd)

    # Consolidated: exact copy of master + added columns
    cons = wb.create_sheet("Consolidated")
    for r in range(1, last + 1):
        for c in range(1, 4):
            s, d = master.cell(r, c), cons.cell(r, c)
            d.value = s.value
            if isinstance(s.value, str) and s.data_type != "f" and d.data_type == "f":
                d.data_type = "s"  # literal text in the master stays literal text
            if s.has_style:
                d.font, d.fill, d.border = copy(s.font), copy(s.fill), copy(s.border)
                d.alignment, d.number_format = copy(s.alignment), s.number_format
    for rng in master.merged_cells.ranges:
        cons.merge_cells(str(rng))
    for c in range(1, 4):
        cons.column_dimensions[L(c)].width = master.column_dimensions[L(c)].width or 30
    adds = DAYS + ["Total planned", "Balance"]
    for j, h in enumerate(adds):
        c = 4 + j
        cell = cons.cell(hdr, c, h)
        src_cell = master.cell(hdr, 3)
        cell.font = copy(src_cell.font) if src_cell.has_style else Font(name=FONT, bold=True)
        cell.fill, cell.border = copy(src_cell.fill), copy(src_cell.border)
        cell.alignment = Alignment(horizontal="center")
        cons.column_dimensions[L(c)].width = 13 if j < 7 else 13
    datarows = [r for r, *_ in rows]
    for r in datarows:
        for j, day in enumerate(DAYS):
            cons.cell(
                r,
                4 + j,
                "=SUMPRODUCT(({d}!$A${f}:$A${l}=$B{r})*({d}!$B${f}:$B${l}=$A{r})*{d}!$E${f}:$E${l})".format(
                    d=day, f=first, l=lastd, r=r
                ),
            )
        cons.cell(r, 10, "=SUM(D{r}:I{r})".format(r=r))
        cons.cell(r, 11, "=N(C{r})-J{r}".format(r=r))
        for c in range(4, 12):
            cons.cell(r, c).number_format = "#,##0"
            cons.cell(r, c).font = Font(name=FONT, size=10)
    tr = last + 2
    cons.cell(tr, 1, "TOTAL")
    for c in range(3, 12):
        cons.cell(tr, c, "=SUM(%s%d:%s%d)" % (L(c), hdr + 1, L(c), last))
        cons.cell(tr, c).number_format = "#,##0"
    for c in range(1, 12):
        cons.cell(tr, c).font = Font(name=FONT, bold=True)
        cons.cell(tr, c).fill = WC_FILL
    cons.conditional_formatting.add(
        "K%d:K%d" % (hdr + 1, last),
        FormulaRule(
            formula=["K%d<0" % (hdr + 1)],
            fill=PatternFill("solid", bgColor="FF0000", fgColor="FF0000"),
            font=Font(color="FFFFFF", bold=True),
        ),
    )
    cons.freeze_panes = cons.cell(hdr + 1, 1)
    cons.auto_filter.ref = "A%d:K%d" % (hdr, last)

    # order: master, days, consolidated, _Lists (hidden)
    order = [mname] + DAYS + ["Consolidated", "_Lists"]
    wb._sheets = [wb[n] for n in order]
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(dst)
    return BuildInfo(
        wcs=wcs,
        items=items,
        first=first,
        lastd=lastd,
        hdr=hdr,
        last=last,
        nrows=nrows,
        rows=rows,
        master_title=mname,
    )
