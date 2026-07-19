"""Dashboard -- the workbook's home screen (PLAN.md §4), built last since
it aggregates almost every other sheet. This module builds the
*structure* only (Phase 2): the KPI strip, five sub-tables, and two
embedded charts, all formula-driven. VBA scope for this sheet is minimal
(§4's Visual/UX note) -- a Workbook_Open full-recalc and a "Refresh Now"
button wired to it -- both Phase 3.

Layout:
  Rows 1-3: KPI strip, 6 cards (§4.A) -- Active Tasks, Completed Tasks
    (Total), Total Hours Invested, Overdue Tasks (red fill when >0),
    Current Campaign Ends, Next Campaign Starts. Same 2-column-per-card
    convention as Monthly Media Plan's §17.10 KPI cards.
  Row 4: Refresh Now button (styled only, Phase 3 wires it).
  Row 6+: left column (A-F) -- Upcoming Tasks (§4.B, top 5) then Yearly
    Monthly Performance (§4.C, 12 rows) with two embedded line charts
    beneath it.
  Row 6+: right column (H-L) -- Daily & Weekly Productivity (§4.D),
    Marketing Campaign Reminders (§4.E, top 10), Campaign Metrics Summary
    (§4.F). Each sub-table is a genuine (fixed-size) Excel Table for
    banding + structured references, per §4's Visual/UX note.

**Two array-formula techniques used to stay CSE-free** (this project
never uses Ctrl+Shift+Enter):
- The two countdown KPIs (§4.A) need "nearest Active/Planned campaign"
  across Campaigns' 20 non-contiguous main rows.

  **Live-tested finding (corrects an initial approach):** the first
  attempt used `AGGREGATE(15,6,array,1)` (SMALL with error-ignore,
  natively array-capable without CSE) dividing by the status condition so
  non-matching rows produce #DIV/0! and get skipped. This computed
  correctly when entered fresh via VBA, but an *openpyxl-written* copy of
  the exact same formula silently fell through to its `IFERROR` fallback
  on a clean file open + full recalc (`Application.Calculate` and even
  `Application.CalculateFullRebuild` both left it wrong) -- confirmed via
  COM that only actually rewriting the cell's formula through Excel (not
  just recalculating it) fixed it. Something about how openpyxl serializes
  this specific array-context AGGREGATE usage doesn't carry the internal
  state real Excel expects, and relying on Phase 3's `Workbook_Open`
  recalc to silently "fix" every such cell on first open would be fragile
  and easy to forget. **Fix:** a `SUMPRODUCT(MIN(...))`-based MIN
  emulation instead -- `SUMPRODUCT` is already used extensively elsewhere
  in this project and reliably forces array-context evaluation on a
  plain openpyxl-written formula. Non-matching rows are pushed to a huge
  sentinel value (`+9999999`, added instead of using `IF()`) so they never
  win the `MIN`, rather than divided into an error -- which in turn means
  `IFERROR` can no longer detect "no matching campaign at all" the way it
  could with `AGGREGATE` (there's no error to catch; every non-match just
  contributes the same huge number). Caught this before it shipped:
  `IF(result>=9000000,"—",result)` checks against the sentinel directly
  instead, since a genuine match is always a small number (single/double-
  digit days, at most low thousands) and a true no-match result is always
  sentinel-dominated regardless of whether the row's own date happened to
  be blank.

  A second, separate live-tested bug in the same formula: it originally
  subtracted `TODAY()` straight from Campaigns' raw Start/End Date columns
  (C/D). Column D is safely blank on every summary row, but column C is
  *not* -- summary rows reuse it for a real `IFERROR(...,"")`-driven
  lookup (campaigns.py's own docstring), so it holds genuine empty-string
  *text* there, not a blank cell. Arithmetic on that text throws
  `#VALUE!`, which -- unlike `AGGREGATE`'s error-ignore option -- neither
  `SUMPRODUCT` nor `MIN` can skip past; the error poisons the whole array
  result. Tried coercing it away with `N(range)` and then
  `IF(ISNUMBER(range),range,0)` inside the `SUMPRODUCT(MIN(...))` array
  context -- both silently collapsed to evaluating only the range's
  *first cell* instead of vectorizing per-element (confirmed via COM: the
  results matched what the top-left cell alone would produce, not the
  correct minimum), unlike bare arithmetic operators on the same ranges,
  which vectorize correctly under `SUMPRODUCT`. **Fix:** reuse
  `ReminderDaysLeft` (Y, added for §4.E below) instead of raw C/D --
  it's already "days to next boundary" for both Active and Planned
  campaigns, and, being genuinely blank (not text) on every summary row,
  needs no coercion at all. Live-verified correct on a freshly-built file
  (never touched by VBA) after only `Application.Calculate`, for both the
  matching and the no-match case.
- Upcoming Tasks (§4.B) and Marketing Campaign Reminders (§4.E) both use
  the established SortKey + SMALL + INDEX/MATCH top-N pattern (Tasks'
  own `SortKey`, and a new `ReminderSortKey` added to Campaigns for this
  --see campaigns.py's docstring). SMALL() silently skips blank cells, so
  it works fine even on Campaigns' sparse, non-contiguous main rows
  without needing a contiguous mirror column.

**The embedded chart** (§4.C/D) was specified as one chart plotting both
Completed (a small monthly count) and Hours (a much larger, differently-
scaled number) together -- sharing one axis would flatten one series
against the other, and combining differently-scaled measures on a
dual-axis chart is avoided here on purpose. Built as **two** small
single-axis line charts side by side instead, using §17.2's already-
assigned slot-1/slot-2 colors for Completed/Hours respectively.
"""

import datetime
import math

from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.table import Table

from . import conditional_formatting as cf
from . import design_system as ds

SHEET_NAME = "Dashboard"

_EUR_FORMAT = '#,##0.00" €"'
_USD_FORMAT = '"$"#,##0.00'
_DATE_FORMAT = "DD.MM.YYYY"
_DAYS_FORMAT = "0"
_HOURS_FORMAT = "0.00"


def _friendly_hours_formula(hours_expr):
    """Wraps an hours-denominated numeric expression (no leading "=") in
    the same "Xh Ym" text rendering already established for Tasks' own
    friendly Time Spent column -- round to the nearest minute FIRST, then
    split into hours/minutes, so a value like 1.999h correctly carries
    into "2h 00m" instead of the nonsensical "1h 60m". User-requested
    (PLAN.md §19 Phase 5): every hours-denominated Dashboard cell showed
    a raw locale-decorated decimal ("0,00" on a comma-decimal machine)
    instead of a friendly duration. Unlike Tasks' own fix, a number-
    format-only fix isn't possible here -- Excel's "[h]:mm" duration
    tokens assume the stored value is a fraction of a DAY, not hours, and
    no format code can perform unit conversion -- so this returns a full
    TEXT formula (a leading "="), replacing the cell's numeric value with
    text, not just redecorating its display. Always shows "0h 00m" for a
    zero value (explicitly requested), not blank.
    """
    return (
        f'=TEXT(INT(ROUND(({hours_expr})*60,0)/60),"0")&"h "&'
        f'TEXT(MOD(ROUND(({hours_expr})*60,0),60),"00")&"m"'
    )

# --- Cross-sheet ranges this sheet reads from ---
_TASKS_ID = "Tasks!$A$2:$A$501"
_TASKS_TASK = "Tasks!$E$2:$E$501"
_TASKS_PROJECT = "Tasks!$H$2:$H$501"
_TASKS_PRIORITY = "Tasks!$J$2:$J$501"
_TASKS_DEADLINE = "Tasks!$K$2:$K$501"
_TASKS_DAYS_LEFT = "Tasks!$M$2:$M$501"
_TASKS_TIME_SPENT = "Tasks!$O$2:$O$501"
_TASKS_SORT_KEY = "Tasks!$W$2:$W$501"

_CAMPAIGNS_NAME = "Campaigns!$C$3:$C$202"
_CAMPAIGNS_START = "Campaigns!$E$3:$E$202"
_CAMPAIGNS_END = "Campaigns!$F$3:$F$202"
_CAMPAIGNS_STATUS = "Campaigns!$G$3:$G$202"
_CAMPAIGNS_REMINDER_DAYS_LEFT = "Campaigns!$AA$3:$AA$202"
_CAMPAIGNS_REMINDER_SORT_KEY = "Campaigns!$AB$3:$AB$202"

_TIME_LOG_DATE = "'Time Log'!$B:$B"
_TIME_LOG_DURATION = "'Time Log'!$C:$C"

_MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

_UPCOMING_TASKS_N = 5
_REMINDERS_N = 10


def _style_header_row(ws, row, first_col, headers):
    ws.row_dimensions[row].height = ds.HEADER_ROW_HEIGHT
    for i, label in enumerate(headers):
        cell = ws.cell(row=row, column=first_col + i, value=label)
        ds.style_header_cell(cell)


def _section_title(ws, cell_ref, text):
    cell = ws[cell_ref]
    cell.value = text
    cell.font = Font(name=ds.FONT_NAME, size=ds.FONT_SIZE_TITLE, bold=True, color=ds.INK_PRIMARY)


def _add_table(ws, name, ref):
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = ds.modern_table_style()
    ws.add_table(table)

    # Quiet row banding + hairline row separators (§17.10), for every
    # Dashboard table via this one shared helper -- min_col/min_row/
    # max_col/max_row parsed straight from `ref` rather than threading
    # extra parameters through every one of this function's 5 call sites.
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    first_data_row = min_row + 1  # skip the header row
    if first_data_row <= max_row:
        col_range = f"{get_column_letter(min_col)}{first_data_row}:{get_column_letter(max_col)}{max_row}"
        cf.apply_rules(ws, cf.row_banding_rule(col_range, first_data_row))
        ds.apply_table_hairline_borders(ws, min_col, max_col, first_data_row, max_row)


def _panel(ws, cell_range):
    """Gives one Dashboard section (title through its last data row) a
    light `CARD_SURFACE` interior and a hairline frame, so it reads as a
    distinct panel sitting on the sheet's own tinted canvas rather than a
    table drawn directly on the page background (§17.16, user-requested:
    "make it feel more like a dashboard, rather than Excel spreadsheet").

    Must be called *before* that section builds its own header band/
    banding/KPI-card fills -- those are more specific and correctly
    paint over this base fill within their own cells; called after, this
    would instead overwrite them. `row_banding_rule()`'s alternating
    stripe already reuses `PAGE_BG`, one shade darker than
    `CARD_SURFACE` -- painting this panel fill first is what makes that
    banding read as *alternating* rather than disappearing into a
    uniform canvas tint (see `build()`'s own canvas-tint comment for why
    the canvas itself uses a third, more visibly distinct shade).
    """
    fill = PatternFill("solid", fgColor=ds.CARD_SURFACE)
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            ws.cell(row=row, column=col).fill = fill
    ds.apply_box_outline(ws, cell_range, color=ds.GRIDLINE)


def build(wb):
    """Create the Dashboard sheet in `wb` (an openpyxl Workbook). Requires
    Tasks, Archive, Campaigns, Campaigns Archive, Monthly Media Plan, and
    Time Log to already exist. Returns the created worksheet.
    """
    ws = wb.create_sheet(SHEET_NAME)
    ds.apply_modern_chrome(ws)
    ws.freeze_panes = "A6"

    # §17.16 originally tinted the whole canvas with `GRIDLINE` here so
    # panels (below) would read as distinct surfaces sitting on top of it.
    # User feedback: didn't like the tinted background, wanted the plain
    # white sheet back -- reverted. The panels' own `CARD_SURFACE` fill and
    # hairline frame (see `_panel()`) still give each section a visible
    # boundary on plain white, so the "framed card" look is unaffected.

    # A bit more vertical breathing room (§17.16) between the KPI/Refresh
    # Now strip and the panels below -- row 5 is already a genuine gap row
    # (nothing else writes into it), so widening it is risk-free, unlike
    # touching any row shared with another section's own title/data.
    ws.row_dimensions[5].height = 10

    _build_kpi_strip(ws)
    _build_refresh_button(ws)
    _build_upcoming_tasks(ws)
    _build_yearly_monthly_performance(ws)
    _build_daily_weekly_productivity(ws)
    _build_campaign_reminders(ws)
    _build_campaign_metrics_summary(ws)

    ws.sheet_properties.tabColor = ds.TAB_CORE
    return ws


# --- A. KPI strip ------------------------------------------------------

_KPI_CARDS = [
    ("Active Tasks (Upcoming)", "A", f"=COUNTA({_TASKS_ID})", "General"),
    ("Completed Tasks (Total)", "D", "=COUNTA(ArchiveTable[ID])", "General"),
    (
        "Total Hours Invested",
        "G",
        _friendly_hours_formula("SUM(Tasks!$O$2:$O$501)+SUM(ArchiveTable[Time Spent (hrs)])"),
        "General",
    ),
    ("Overdue Tasks", "J", f'=COUNTIFS({_TASKS_DAYS_LEFT},"<0")', "General"),
    (
        "Current Campaign Ends (days)",
        "M",
        (
            f'=IF(SUMPRODUCT(MIN({_CAMPAIGNS_REMINDER_DAYS_LEFT}+({_CAMPAIGNS_STATUS}<>"Active")*9999999))>=9000000,"—",'
            f'SUMPRODUCT(MIN({_CAMPAIGNS_REMINDER_DAYS_LEFT}+({_CAMPAIGNS_STATUS}<>"Active")*9999999)))'
        ),
        _DAYS_FORMAT,
    ),
    (
        "Next Campaign Starts (days)",
        "P",
        (
            f'=IF(SUMPRODUCT(MIN({_CAMPAIGNS_REMINDER_DAYS_LEFT}+({_CAMPAIGNS_STATUS}<>"Planned")*9999999))>=9000000,"—",'
            f'SUMPRODUCT(MIN({_CAMPAIGNS_REMINDER_DAYS_LEFT}+({_CAMPAIGNS_STATUS}<>"Planned")*9999999)))'
        ),
        _DAYS_FORMAT,
    ),
]


def _build_kpi_strip(ws):
    # Modernization pass (§17.10, PLAN.md §19 Phase 6, user-requested):
    # a thick brand navy top accent bar reads as a more current "stat
    # tile" pattern than the original left-edge color bar (a heavier,
    # more dated affordance) -- the rest of the card outline is now a
    # plain hairline GRIDLINE box, so the accent color is the one clear
    # signal, not competing with itself on multiple edges.
    for caption, col_letter, formula, number_format in _KPI_CARDS:
        first_col_idx = get_column_letter_index(col_letter)
        next_col = get_column_letter(first_col_idx + 1)

        caption_ref = f"{col_letter}1"
        ws.merge_cells(f"{caption_ref}:{next_col}1")
        caption_cell = ws[caption_ref]
        caption_cell.value = caption
        caption_cell.font = Font(name=ds.FONT_NAME, size=ds.FONT_SIZE_CAPTION, color=ds.INK_SECONDARY)
        # wrap=True + a taller row 1 (§17.11 correction, user-reported
        # crop on "Completed Tasks (Total)") -- a caption this long at
        # 9pt in a 2-column-wide card was clipping instead of wrapping.
        caption_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 28

        value_ref = f"{col_letter}2"
        ws.merge_cells(f"{value_ref}:{next_col}3")
        value_cell = ws[value_ref]
        value_cell.value = formula
        value_cell.number_format = number_format
        value_cell.font = Font(name=ds.FONT_NAME, size=ds.FONT_SIZE_KPI, bold=True, color=ds.INK_PRIMARY)
        value_cell.alignment = Alignment(horizontal="center", vertical="center")

        card_fill = PatternFill("solid", fgColor=ds.CARD_SURFACE)
        hair = Side(style="hair", color=ds.GRIDLINE)
        thick_top = Side(style="thick", color=ds.BRAND_ACCENT)
        # Every physical cell across all 3 rows/2 cols of the card, not
        # just the two merges' anchor cells (§17.10 bugfix -- a border set
        # only on a merge's anchor cell leaves the far column/row edge
        # unbordered in real Excel; see design_system.apply_box_outline's
        # docstring). Row 1 is its own merge (thick blue top accent, per
        # this pass's KPI redesign), rows 2-3 are a second merge (the
        # value) -- handled together here rather than via that shared
        # helper since the top accent's style/color genuinely differs
        # from the rest of the box, which apply_box_outline assumes is
        # uniform.
        for row in (1, 2, 3):
            for c in (first_col_idx, first_col_idx + 1):
                cell = ws.cell(row=row, column=c)
                cell.fill = card_fill
                cell.border = Border(
                    top=thick_top if row == 1 else None,
                    bottom=hair if row == 3 else None,
                    left=hair if c == first_col_idx else None,
                    right=hair if c == first_col_idx + 1 else None,
                )

        for c in range(first_col_idx, first_col_idx + 2):
            ws.column_dimensions[get_column_letter(c)].width = 13

    # Overdue Tasks: red fill when > 0 (§4.A) -- bgColor, not fgColor, see
    # conditional_formatting.py's _fill() docstring.
    ws.conditional_formatting.add(
        "J2:K3",
        FormulaRule(formula=["$J$2>0"], fill=PatternFill("solid", bgColor=ds.STATUS_CRITICAL)),
    )

    # Current Campaign Ends (days): a status indicator on the card's own
    # top accent bar (§17.16, user-requested "trend/status indicators
    # next to KPI numbers") -- amber when a currently-Active campaign
    # ends within 2 days, otherwise the default brand navy stays
    # (matches the deliberate "one-state alert" design of the Overdue
    # Tasks fill above: color signals only when something needs
    # attention, not a 3-state scheme competing for the eye on every
    # card). Same CF-driven border technique as Calendar's own selected-
    # date/today border rules (`FormulaRule(..., border=...)`), applied
    # to row 1 specifically since that's where the top accent bar lives.
    # ISNUMBER-guarded -- this KPI shows "—" (text) when no Active
    # campaign exists at all, and bare arithmetic on that would error.
    # All 4 sides explicitly set, not just `top` -- confirmed via COM that
    # a differential-style (CF) Border with only one side set doesn't
    # serialize correctly (the CF rule existed but its own Top border
    # threw "cannot read LineStyle" when queried back, meaning it was
    # never actually written); Calendar's own working CF borders
    # (selected-date/today) already set all 4 sides uniformly, unlike
    # this one's original attempt.
    warning_hair = Side(style="hair", color=ds.GRIDLINE)
    warning_top_border = Border(
        top=Side(style="thick", color=ds.STATUS_WARNING),
        bottom=warning_hair,
        left=warning_hair,
        right=warning_hair,
    )
    ws.conditional_formatting.add(
        "M1:N1",
        FormulaRule(formula=["AND(ISNUMBER($M$2),$M$2<=2)"], border=warning_top_border),
    )


def get_column_letter_index(letter):
    from openpyxl.utils import column_index_from_string

    return column_index_from_string(letter)


def _build_refresh_button(ws):
    ws.merge_cells("A4:C4")
    button = ws.cell(row=4, column=1, value=f"{ds.ICON_REFRESH} Refresh Now")
    button.font = Font(name=ds.FONT_NAME, size=ds.FONT_SIZE_BODY, bold=True, color=ds.INK_PRIMARY)
    button.fill = PatternFill("solid", fgColor=ds.CARD_SURFACE)
    ds.apply_box_outline(ws, "A4:C4", color=ds.INK_MUTED)
    button.alignment = Alignment(horizontal="center", vertical="center")


# --- B. Upcoming Tasks (top 5) -----------------------------------------

def _build_upcoming_tasks(ws):
    _panel(ws, "A6:F12")  # §17.16 -- last_row = first_row(8) + N(5) - 1
    _section_title(ws, "A6", "Upcoming Tasks (Next 5 Deadlines)")
    headers = ["ID", "Task", "Project", "Priority", "Deadline", "Days Left"]
    _style_header_row(ws, 7, 1, headers)

    sources = [_TASKS_ID, _TASKS_TASK, _TASKS_PROJECT, _TASKS_PRIORITY, _TASKS_DEADLINE, _TASKS_DAYS_LEFT]
    first_row = 8
    for k in range(1, _UPCOMING_TASKS_N + 1):
        row = first_row + k - 1
        for col_offset, source in enumerate(sources, start=1):
            cell = ws.cell(row=row, column=col_offset)
            cell.value = (
                f"=IFERROR(INDEX({source},MATCH(SMALL({_TASKS_SORT_KEY},{k}),{_TASKS_SORT_KEY},0)),\"\")"
            )
        ws.cell(row=row, column=5).number_format = _DATE_FORMAT
        ws.cell(row=row, column=6).number_format = _DAYS_FORMAT
        # §17.11 correction (user-reported): center ID and Days Left.
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row, column=6).alignment = Alignment(horizontal="center", vertical="center")

    last_row = first_row + _UPCOMING_TASKS_N - 1
    for col in ("A", "B", "C", "D", "E", "F"):
        ws.column_dimensions[col].width = 12
    ws.column_dimensions["B"].width = 26

    # Highlight rows with Days Left <= 2 (§4.B), across the whole row block.
    ws.conditional_formatting.add(
        f"A{first_row}:F{last_row}",
        FormulaRule(formula=[f'AND($F{first_row}<>"",$F{first_row}<=2)'], fill=PatternFill("solid", bgColor=ds.STATUS_WARNING)),
    )

    _add_table(ws, "UpcomingTasksTable", f"A7:F{last_row}")


# --- C. Yearly Monthly Performance + charts -----------------------------

def _build_yearly_monthly_performance(ws):
    title_row = 14
    header_row = 15
    first_row = 16
    last_row = first_row + 11  # 12 months

    _panel(ws, f"A{title_row}:C{last_row}")  # §17.16 -- D is the hidden HoursRaw helper, excluded
    _section_title(ws, f"A{title_row}", "Yearly Monthly Performance")
    _style_header_row(ws, header_row, 1, ["Month", "Completed", "Hours"])
    # "HoursRaw" (D) -- a raw-numeric helper, deliberately NOT part of the
    # visible table (ref stays A:C below) and hidden from the sheet. User-
    # requested (PLAN.md §19 Phase 5): "Hours" (C) itself needed to become
    # a friendly "Xh Ym" TEXT cell like every other hours display on this
    # sheet, but this column also feeds the "Hours Invested by Month"
    # chart below, which needs an actual NUMBER to plot -- a number format
    # can't produce "Xh Ym" for an hours-stored value (no unit conversion
    # in format codes), and a text cell can't drive a line chart. Splitting
    # the number and its friendly display into two cells, chart pointed at
    # the hidden one, is the same resolution used for Tasks' own friendly
    # Time Spent column (raw hrs hidden, friendly display added
    # separately) -- see _build_performance_charts' visible_cells_only
    # override, needed because Excel charts skip hidden data by default.
    ws.cell(row=header_row, column=4, value="HoursRaw")
    ws.column_dimensions["D"].hidden = True

    for m in range(1, 13):
        row = first_row + m - 1
        ws.cell(row=row, column=1, value=_MONTH_NAMES[m - 1])

        next_month_start = f"DATE(YEAR(TODAY()),{m + 1},1)" if m < 12 else "DATE(YEAR(TODAY())+1,1,1)"
        month_start = f"DATE(YEAR(TODAY()),{m},1)"

        center = Alignment(horizontal="center", vertical="center")

        completed_cell = ws.cell(row=row, column=2)
        completed_cell.value = (
            f'=COUNTIFS(ArchiveTable[Completed At],">="&{month_start},'
            f'ArchiveTable[Completed At],"<"&{next_month_start})'
        )
        completed_cell.alignment = center  # §17.11 correction (user-reported)

        hours_raw_expr = (
            f'SUMIFS(ArchiveTable[Time Spent (hrs)],ArchiveTable[Completed At],">="&{month_start},'
            f'ArchiveTable[Completed At],"<"&{next_month_start})'
        )
        hours_raw_cell = ws.cell(row=row, column=4)
        hours_raw_cell.value = f"={hours_raw_expr}"
        hours_raw_cell.number_format = _HOURS_FORMAT

        hours_cell = ws.cell(row=row, column=3)
        hours_cell.value = _friendly_hours_formula(hours_raw_expr)
        hours_cell.alignment = center  # §17.11 correction (user-reported)

    ws.column_dimensions["A"].width = 12

    _add_table(ws, "YearlyMonthlyPerformanceTable", f"A{header_row}:C{last_row}")

    _build_performance_charts(ws, header_row, first_row, last_row)


def _build_performance_charts(ws, header_row, first_row, last_row):
    # §17.10 bugfix (user-reported live): these two charts used to sit
    # side-by-side (anchored at columns A and E), which silently depended
    # on columns A-D being narrower, combined, than each chart's own 10cm
    # width -- untrue on this sheet (Dashboard's column widths are set and
    # re-set by several unrelated sections throughout build()), so the
    # charts overlapped each other by a sliver AND the second chart
    # extended into column H (~65% of its width) covering most of
    # Campaign Metrics Summary, which lives there. Fixed by stacking both
    # charts vertically in column A instead -- removes the dependency on
    # any other section's column widths entirely, rather than hand-tuning
    # anchor columns/widths to just barely avoid the same class of bug
    # recurring the next time a column width changes elsewhere.
    chart_anchor_row = last_row + 2
    row_height_pt = ws.sheet_format.defaultRowHeight or 15
    chart_height_cm = 6
    # generous: a 6cm-tall chart, converted to rows via each row's own
    # point height (1pt = 1/72in = 2.54/72cm), plus a couple of rows of
    # visible gap -- rounded up, not tuned tight, so this stays correct
    # even if the sheet's default row height changes again later.
    rows_per_chart = math.ceil((chart_height_cm * 72 / 2.54) / row_height_pt) + 2

    completed_chart = LineChart()
    completed_chart.title = "Completed Tasks by Month"
    completed_chart.style = None
    completed_chart.height = chart_height_cm
    completed_chart.width = 10
    data = Reference(ws, min_col=2, max_col=2, min_row=header_row, max_row=last_row)
    cats = Reference(ws, min_col=1, min_row=first_row, max_row=last_row)
    completed_chart.add_data(data, titles_from_data=True)
    completed_chart.set_categories(cats)
    completed_chart.series[0].graphicalProperties.line.solidFill = ds.CATEGORICAL[0]
    completed_chart.series[0].smooth = False
    ws.add_chart(completed_chart, f"A{chart_anchor_row}")

    hours_chart = LineChart()
    hours_chart.title = "Hours Invested by Month"
    hours_chart.style = None
    hours_chart.height = chart_height_cm
    hours_chart.width = 10
    # Column D (HoursRaw), not C ("Hours") -- C is now a friendly "Xh Ym"
    # TEXT cell (user-requested), which a line chart can't plot; D holds
    # the same underlying number, hidden from the sheet. visible_cells_only
    # = False (Excel's own default is True -- charts skip hidden data
    # otherwise, confirmed via openpyxl's chartspace.py plotVisOnly flag)
    # so the chart still plots D's values despite the column being hidden.
    hours_chart.visible_cells_only = False
    data2 = Reference(ws, min_col=4, max_col=4, min_row=header_row, max_row=last_row)
    hours_chart.add_data(data2, titles_from_data=True)
    hours_chart.set_categories(cats)
    hours_chart.series[0].graphicalProperties.line.solidFill = ds.CATEGORICAL[1]
    hours_chart.series[0].smooth = False
    ws.add_chart(hours_chart, f"A{chart_anchor_row + rows_per_chart}")


# --- D. Daily & Weekly Productivity -------------------------------------

def _build_daily_weekly_productivity(ws):
    title_row = 6
    header_row = 7
    first_row = 8

    _panel(ws, "H6:I11")  # §17.16 -- last_row = first_row(8) + 4 rows - 1
    _section_title(ws, f"H{title_row}", "Daily & Weekly Productivity")
    _style_header_row(ws, header_row, 8, ["Productivity Indicator", "Performance"])

    week_start = "(TODAY()-WEEKDAY(TODAY(),3))"
    week_end = f"({week_start}+7)"

    rows = [
        ("Completed Tasks Today", '=COUNTIFS(ArchiveTable[Completed At],">="&TODAY(),ArchiveTable[Completed At],"<"&TODAY()+1)', "General"),
        (
            "Time Invested Today",
            _friendly_hours_formula(f"SUMIFS({_TIME_LOG_DURATION},{_TIME_LOG_DATE},TODAY())"),
            "General",
        ),
        (
            "Completed This Week",
            f'=COUNTIFS(ArchiveTable[Completed At],">="&{week_start},ArchiveTable[Completed At],"<"&{week_end})',
            "General",
        ),
        (
            "Time Invested This Week",
            _friendly_hours_formula(
                f'SUMIFS({_TIME_LOG_DURATION},{_TIME_LOG_DATE},">="&{week_start},{_TIME_LOG_DATE},"<"&{week_end})'
            ),
            "General",
        ),
    ]
    for i, (label, formula, number_format) in enumerate(rows):
        row = first_row + i
        ws.cell(row=row, column=8, value=label)
        value_cell = ws.cell(row=row, column=9, value=formula)
        value_cell.number_format = number_format
        value_cell.alignment = Alignment(horizontal="center", vertical="center")  # §17.11 (user-reported)

    last_row = first_row + len(rows) - 1
    # 28, not 22 (§17.11 correction, user-reported crop) -- this column is
    # shared with Campaign Metrics Summary below it (both live in H:I),
    # and that section's longest label ("Planned Budget Total (EUR)", 27
    # characters) was clipping against column I's non-empty value cells
    # at the old width.
    ws.column_dimensions["H"].width = 28
    ws.column_dimensions["I"].width = 14

    _add_table(ws, "DailyWeeklyProductivityTable", f"H{header_row}:I{last_row}")


# --- E. Marketing Campaign Reminders (7-day window, top 10) -------------

def _build_campaign_reminders(ws):
    title_row = 13
    header_row = 14
    first_row = 15
    last_row = first_row + _REMINDERS_N - 1

    _panel(ws, f"H{title_row}:L{last_row}")  # §17.16
    _section_title(ws, f"H{title_row}", "Marketing Campaign Reminders (7-Day Window)")
    _style_header_row(ws, header_row, 8, ["Campaign Name", "Start Date", "End Date", "Days Left", "Status"])

    sources = [_CAMPAIGNS_NAME, _CAMPAIGNS_START, _CAMPAIGNS_END, _CAMPAIGNS_REMINDER_DAYS_LEFT, _CAMPAIGNS_STATUS]
    for k in range(1, _REMINDERS_N + 1):
        row = first_row + k - 1
        for col_offset, source in enumerate(sources, start=8):
            cell = ws.cell(row=row, column=col_offset)
            cell.value = (
                f"=IFERROR(INDEX({source},MATCH(SMALL({_CAMPAIGNS_REMINDER_SORT_KEY},{k}),"
                f"{_CAMPAIGNS_REMINDER_SORT_KEY},0)),\"\")"
            )
        ws.cell(row=row, column=9).number_format = _DATE_FORMAT
        ws.cell(row=row, column=10).number_format = _DATE_FORMAT
        ws.cell(row=row, column=11).number_format = _DAYS_FORMAT
        ws.cell(row=row, column=11).alignment = Alignment(horizontal="center", vertical="center")  # §17.11

    ws.column_dimensions["J"].width = 13
    ws.column_dimensions["K"].width = 13
    ws.column_dimensions["L"].width = 11

    _add_table(ws, "MarketingCampaignRemindersTable", f"H{header_row}:L{last_row}")


# --- F. Campaign Metrics Summary ----------------------------------------

def _build_campaign_metrics_summary(ws):
    title_row = 26
    header_row = 27
    first_row = 28

    _panel(ws, f"H{title_row}:I32")  # §17.16 -- last_row = first_row(28) + 5 rows - 1
    _section_title(ws, f"H{title_row}", "Campaign Metrics Summary")
    _style_header_row(ws, header_row, 8, ["Performance Metric", "Count / Total"])

    rows = [
        ("Active Campaigns Count", f'=COUNTIFS({_CAMPAIGNS_STATUS},"Active")', "General"),
        ("Planned Campaigns Count", f'=COUNTIFS({_CAMPAIGNS_STATUS},"Planned")', "General"),
        ("Completed Campaigns Count", "=COUNTA(CampaignsArchiveTable[Campaign Name])", "General"),
        ("Planned Budget Total (EUR)", "=TotalPlannedBudgetEUR", _EUR_FORMAT),
        ("Planned Budget Total (USD)", "=TotalPlannedBudgetUSD", _USD_FORMAT),
    ]
    for i, (label, formula, number_format) in enumerate(rows):
        row = first_row + i
        ws.cell(row=row, column=8, value=label)
        value_cell = ws.cell(row=row, column=9, value=formula)
        value_cell.number_format = number_format
        value_cell.alignment = Alignment(horizontal="center", vertical="center")  # §17.11 (user-reported)

    last_row = first_row + len(rows) - 1
    _add_table(ws, "CampaignMetricsSummaryTable", f"H{header_row}:I{last_row}")
