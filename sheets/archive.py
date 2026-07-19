"""Archive -- where completed task branches land automatically (PLAN.md
§7). This module builds the *structure* only (Phase 2): columns,
formatting, conditional formatting, sheet protection. Unlike Tasks,
nothing here is a live formula -- every value is a static snapshot
written by the archiving macro at archive time (Phase 3, see
vba/modHierarchy.bas's companion archiving routine).

Column layout (§7.1), left to right:
  ID, Start Date, Task, Total Subtasks, Subtask of, Project/Category,
  Linked Campaign, Priority, Deadline, Status, Days Left, Time Spent,
  Notes, Completed At
Hidden columns, grouped together after Completed At:
  TaskKey, ParentKey, Level

`Archive` is a genuine auto-expanding Excel Table, like Time Log --
unlike Tasks (whose rows get wholesale rewritten in place by
RenumberHierarchy), Archive rows are append-only and never reordered
(§7.2), so it starts at the minimum valid size and grows via VBA's
ListRows.Add.

The whole sheet is protected (locked against direct cell edits) per the
confirmed decision in §7.1's correction note -- Dashboard's historical
stats read straight from this data. Only basic protection is set here;
`UserInterfaceOnly:=True` (letting VBA keep writing while blocking direct
typing) does not persist across file close/reopen and must be re-asserted
by Phase 3's Workbook_Open, or the archiving macro's own writes would be
blocked too.
"""

from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table

from . import conditional_formatting as cf
from . import design_system as ds

SHEET_NAME = "Archive"
HEADER_ROW = 1
FIRST_DATA_ROW = 2

_DATE_TIME_FORMAT = "DD.MM.YYYY HH:MM"
_HOURS_FORMAT = "0.00"
_DAYS_FORMAT = "0"

HEADERS = [
    "ID",
    "Start Date",
    "Task",
    "Total Subtasks",
    "Subtask of",
    "Project / Category",
    "Linked Campaign",
    "Priority",
    "Deadline",
    "Status",
    "Days Left",
    "Time Spent (hrs)",
    "Notes",
    "Completed At",
]
HIDDEN_HEADERS = ["TaskKey", "ParentKey", "Level"]

# Friendly "Xh Ym" Time Spent display (§17.11 correction, PLAN.md §19
# Phase 6, user-reported): "Time Spent" used to be the raw hours number
# itself (a plain "0.00"-formatted decimal, same complaint the rest of
# the workbook already had before its own friendly-hours fix). Archive
# has no live formulas elsewhere -- every value here is a frozen snapshot
# written once by the archiving macro -- but a formula that only
# references its OWN row's OWN "Time Spent (hrs)" cell is still safe to
# add: it's position-relative, not dependent on anything outside that
# row, the same reasoning Tasks' own friendly Time Spent column already
# relies on. Appended after the hidden hierarchy-key block (not next to
# "Time Spent (hrs)" itself) for the same reason Tasks appended its own
# friendly column at the far end -- avoids renumbering every downstream
# reference to this sheet's column layout. "Time Spent (hrs)" itself is
# now hidden (see build()); Dashboard's two `ArchiveTable[Time Spent]`
# formulas were updated to `ArchiveTable[Time Spent (hrs)]` to match.
_FRIENDLY_TIME_SPENT_COL = len(HEADERS) + len(HIDDEN_HEADERS) + 1

_COL = {name: i + 1 for i, name in enumerate(HEADERS)}
_HIDDEN_COL = {name: len(HEADERS) + i + 1 for i, name in enumerate(HIDDEN_HEADERS)}


def build(wb):
    """Create the Archive sheet in `wb` (an openpyxl Workbook). Returns
    the created worksheet.
    """
    ws = wb.create_sheet(SHEET_NAME)
    ds.apply_modern_chrome(ws)
    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=1)
    ws.row_dimensions[HEADER_ROW].height = ds.HEADER_ROW_HEIGHT

    for col, label in enumerate([*HEADERS, *HIDDEN_HEADERS], start=1):
        cell = ws.cell(row=HEADER_ROW, column=col, value=label)
        ds.style_header_cell(cell)
    friendly_ts_cell = ws.cell(row=HEADER_ROW, column=_FRIENDLY_TIME_SPENT_COL, value="Time Spent")
    ds.style_header_cell(friendly_ts_cell, divider=False)

    for letter in (
        get_column_letter(_HIDDEN_COL["TaskKey"]),
        get_column_letter(_HIDDEN_COL["ParentKey"]),
        get_column_letter(_HIDDEN_COL["Level"]),
        get_column_letter(_COL["Time Spent (hrs)"]),
    ):
        ws.column_dimensions[letter].hidden = True

    # Start Date/Deadline/Completed At: 19, not 15 (§17.10 bugfix, same
    # zoom-vs-column-width regression as Tasks -- see its own note).
    widths = {
        "ID": 10,
        "Start Date": 19,
        "Task": 32,
        "Total Subtasks": 10,
        "Subtask of": 10,
        "Project / Category": 18,
        "Linked Campaign": 18,
        "Priority": 10,
        "Deadline": 19,
        "Status": 13,
        "Days Left": 10,
        "Notes": 8,
        "Completed At": 19,
    }
    for name, width in widths.items():
        ws.column_dimensions[get_column_letter(_COL[name])].width = width
    ws.column_dimensions[get_column_letter(_FRIENDLY_TIME_SPENT_COL)].width = 12

    # Number formats on the initial data row (row 2) -- VBA's ListRows.Add
    # (Phase 3) inherits formatting from the row above as the table grows,
    # same pattern already verified working for Time Log. ID and Subtask
    # of specifically need Text format ("@") for the exact same reason
    # Tasks' own ID/Subtask of columns do (modHierarchy.bas's own
    # comment): Excel silently reinterprets a numeric-looking string like
    # "1.1"/"1.2" as an actual number otherwise -- confirmed live
    # (PLAN.md §19 Phase 5) that the archiving macro's plain VBA .Value =
    # assignment hits this exact bug on Archive's un-formatted General
    # cells, even though the SOURCE cell on Tasks is correctly text-typed.
    ws.cell(row=FIRST_DATA_ROW, column=_COL["ID"]).number_format = "@"
    ws.cell(row=FIRST_DATA_ROW, column=_COL["Subtask of"]).number_format = "@"
    ws.cell(row=FIRST_DATA_ROW, column=_COL["Start Date"]).number_format = _DATE_TIME_FORMAT
    ws.cell(row=FIRST_DATA_ROW, column=_COL["Deadline"]).number_format = _DATE_TIME_FORMAT
    ws.cell(row=FIRST_DATA_ROW, column=_COL["Days Left"]).number_format = _DAYS_FORMAT
    ws.cell(row=FIRST_DATA_ROW, column=_COL["Time Spent (hrs)"]).number_format = _HOURS_FORMAT
    ws.cell(row=FIRST_DATA_ROW, column=_COL["Notes"]).number_format = ds.notes_icon_format()
    ws.cell(row=FIRST_DATA_ROW, column=_COL["Completed At"]).number_format = _DATE_TIME_FORMAT

    # §17.11 correction (user-reported): centered alignment for ID, Days
    # Left, Subtask of, Number of Subtasks -- same "row 2 sets it, VBA's
    # ListRows.Add inherits it" convention as the number formats above.
    center = Alignment(horizontal="center", vertical="center")
    for name in ("ID", "Total Subtasks", "Subtask of", "Days Left"):
        ws.cell(row=FIRST_DATA_ROW, column=_COL[name]).alignment = center

    # Friendly "Xh Ym" Time Spent -- position-relative formula (see this
    # module's own header-block comment for why this is safe despite
    # Archive otherwise being static-value-only), same carry-safe
    # round-to-nearest-minute-first pattern as Tasks' own friendly column.
    ts_ref = f"{get_column_letter(_COL['Time Spent (hrs)'])}{FIRST_DATA_ROW}"
    friendly_ts_value_cell = ws.cell(row=FIRST_DATA_ROW, column=_FRIENDLY_TIME_SPENT_COL)
    friendly_ts_value_cell.value = (
        f'=TEXT(INT(ROUND({ts_ref}*60,0)/60),"0")&"h "&TEXT(MOD(ROUND({ts_ref}*60,0),60),"00")&"m"'
    )
    friendly_ts_value_cell.alignment = center

    last_col_letter = get_column_letter(_FRIENDLY_TIME_SPENT_COL)
    table = Table(displayName="ArchiveTable", ref=f"A1:{last_col_letter}{FIRST_DATA_ROW}")
    table.tableStyleInfo = ds.modern_table_style()
    ws.add_table(table)

    # --- Conditional formatting: Priority chip (§17.3) + diverging Days
    # Left scale (§17.5) -- CF ranges extend generously below the initial
    # table row since the table (and therefore these rules, which apply
    # to a fixed range, not the ListObject itself) will grow over time.
    _CF_LAST_ROW = 1000
    priority_letter = get_column_letter(_COL["Priority"])
    days_left_letter = get_column_letter(_COL["Days Left"])
    cf.apply_rules(ws, cf.priority_rules(priority_letter, FIRST_DATA_ROW, _CF_LAST_ROW))
    cf.apply_rules(ws, cf.diverging_rules(days_left_letter, FIRST_DATA_ROW, _CF_LAST_ROW))
    # Quiet row banding (§17.10) across the same generous growth range --
    # no per-row hairline borders here (unlike the fixed-size 50-row
    # reference sheets): a direct Border object on every cell of a
    # 1000-row range would mean real per-cell file bloat, unlike a CF
    # rule which is stored once regardless of range size. Banding alone
    # already reads cleanly without it.
    cf.apply_rules(ws, cf.row_banding_rule(f"A{FIRST_DATA_ROW}:{last_col_letter}{_CF_LAST_ROW}", FIRST_DATA_ROW))

    # --- Sheet protection: locked against direct edits (confirmed with
    # the user, §7.1 correction). UserInterfaceOnly must be re-asserted
    # by Phase 3's Workbook_Open -- see module docstring.
    ws.protection.sheet = True

    # --- Sheet tab color (archive sheet, §17.9) ---
    ws.sheet_properties.tabColor = ds.TAB_ARCHIVE

    return ws
