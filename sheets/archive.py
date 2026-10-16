"""Archive -- where completed task branches land automatically (PLAN.md
§7). This module builds the *structure* only (Phase 2): columns,
formatting, conditional formatting, sheet protection. Unlike Tasks,
nothing here is a live formula -- every value is a static snapshot
written by the archiving macro at archive time (Phase 3, see
vba/modHierarchy.bas's companion archiving routine).

`build(wb, variant="employee")` supports two layouts sharing this one
module rather than a forked file, per the manager-variant design (PLAN.md,
"Manager-variant workbook" §6): the `variant="employee"` path is
untouched byte-for-byte from before this parameter existed.

Employee column layout (V4, PLAN.md §22.5), left to right:
  Completed At, Task, Project/Category, Linked Campaign, Priority,
  Deadline, Status, Early / Late, Time Spent, Notes, Start Date,
  Subtask of, Total Subtasks
Hidden columns, grouped together after Total Subtasks:
  ID, TaskKey, ParentKey, Level, Time Spent (hrs)
Completed At leads because "when did I do this" is how the archive is
actually read; ID is hidden because it's a tree-position snapshot (almost
always "1" out of context). "Early / Late" is the old "Days Left" (same
frozen Deadline - Completed At value) under an honest name. The friendly
Time Spent display is now a real table column (it used to sit outside
the table and only ever had a formula in its first row, PLAN.md §22.2
#1). VBA writes every column by header name (modTasks.ArchiveHeaderMap),
so this order can change without touching the archiving code.

Manager column layout (mirrors Tasks' own positioning, per plan §6):
  ID, Start Date, Task, PIC, Total Subtasks, Completed Subtasks,
  Subtask of, Project/Category, Linked Campaign, Priority, Deadline,
  Status, Days Left, Notes, Completed At
Hidden columns, grouped together after Completed At:
  TaskKey, ParentKey, Level, RowType
No Time Spent / friendly Time Spent at all (manager tracks no timer or
hours). RowType (hidden) is "Completed" or "Parent Context" -- see
vba/modTasksManager.bas's CascadeComplete/FindOrCreateArchiveRow (Phase
5), which this column layout was already built to feed ahead of this
sheet's own construction (same "logic ahead of structure" approach Phase
4 used for Tasks' own manager layout). A "Parent Context" row is a
still-open ancestor's live snapshot (kept current every time a descendant
completes underneath it, per plan §6) so a completed subtask's whole
ancestor chain stays visible even while those ancestors are still open in
Tasks; it converts to "Completed" in place, never duplicated, the moment
that ancestor itself later completes for real.

`Archive` is a genuine auto-expanding Excel Table in both variants, like
Time Log -- unlike Tasks (whose rows get wholesale rewritten in place by
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

from openpyxl.styles import Alignment, PatternFill
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

EMPLOYEE_HEADERS = [
    "Completed At",
    "Task",
    "Project / Category",
    "Linked Campaign",
    "Priority",
    "Deadline",
    "Status",
    "Early / Late",
    "Time Spent",
    "Notes",
    "Start Date",
    "Subtask of",
    "Total Subtasks",
]
EMPLOYEE_HIDDEN_HEADERS = ["ID", "TaskKey", "ParentKey", "Level", "Time Spent (hrs)"]

# Frozen Deadline - Completed At, in days: whole days either way, anything
# within half a day reads "on time".
_EARLY_LATE_FORMAT = '[>=0.5]0"d early";[<=-0.5]0"d late";"on time"'

MANAGER_HEADERS = [
    "ID",
    "Start Date",
    "Task",
    "PIC",
    "Total Subtasks",
    "Completed Subtasks",
    "Subtask of",
    "Project / Category",
    "Linked Campaign",
    "Priority",
    "Deadline",
    "Status",
    "Days Left",
    "Notes",
    "Completed At",
]
MANAGER_HIDDEN_HEADERS = ["TaskKey", "ParentKey", "Level", "RowType"]

# Kept for backward compatibility with any external reference to the old
# unconditional (employee-only) module-level names.
HEADERS = EMPLOYEE_HEADERS
HIDDEN_HEADERS = EMPLOYEE_HIDDEN_HEADERS



def build(wb, variant="employee"):
    """Create the Archive sheet in `wb` (an openpyxl Workbook). `variant`
    is "employee" (default, unchanged from before this parameter existed)
    or "manager" (PIC + Completed Subtasks + RowType, no Time Spent).
    Returns the created worksheet.
    """
    assert variant in ("employee", "manager"), f"unknown variant: {variant!r}"
    is_manager = variant == "manager"

    headers = MANAGER_HEADERS if is_manager else EMPLOYEE_HEADERS
    hidden_headers = MANAGER_HIDDEN_HEADERS if is_manager else EMPLOYEE_HIDDEN_HEADERS
    col = {name: i + 1 for i, name in enumerate(headers)}
    hidden_col = {name: len(headers) + i + 1 for i, name in enumerate(hidden_headers)}

    ws = wb.create_sheet(SHEET_NAME)
    ds.apply_modern_chrome(ws)
    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=1)
    ws.row_dimensions[HEADER_ROW].height = ds.HEADER_ROW_HEIGHT

    for c, label in enumerate([*headers, *hidden_headers], start=1):
        cell = ws.cell(row=HEADER_ROW, column=c, value=label)
        ds.style_header_cell(cell)

    hidden_letters = [get_column_letter(hidden_col[name]) for name in hidden_headers]
    for letter in hidden_letters:
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
        "Notes": 8,
        "Completed At": 19,
    }
    if is_manager:
        widths["Days Left"] = 10
        widths["PIC"] = 16
        widths["Completed Subtasks"] = 12
    else:
        del widths["ID"]  # hidden for employee
        widths["Early / Late"] = 11
        widths["Time Spent"] = 12
    # Employee's "Time Spent (hrs)" column is hidden and deliberately has
    # no explicit width set here (matches the pre-variant original -- only
    # the visible friendly Time Spent column below gets one).
    for name, width in widths.items():
        ws.column_dimensions[get_column_letter(col[name])].width = width

    # Number formats on the initial data row (row 2) -- VBA's ListRows.Add
    # (Phase 3/5) inherits formatting from the row above as the table
    # grows, same pattern already verified working for Time Log. ID and
    # Subtask of specifically need Text format ("@") for the exact same
    # reason Tasks' own ID/Subtask of columns do (modHierarchy.bas's own
    # comment): Excel silently reinterprets a numeric-looking string like
    # "1.1"/"1.2" as an actual number otherwise -- confirmed live
    # (PLAN.md §19 Phase 5) that the archiving macro's plain VBA .Value =
    # assignment hits this exact bug on Archive's un-formatted General
    # cells, even though the SOURCE cell on Tasks is correctly text-typed.
    all_col = {**col, **hidden_col}
    days_left_name = "Days Left" if is_manager else "Early / Late"
    ws.cell(row=FIRST_DATA_ROW, column=all_col["ID"]).number_format = "@"
    ws.cell(row=FIRST_DATA_ROW, column=col["Subtask of"]).number_format = "@"
    ws.cell(row=FIRST_DATA_ROW, column=col["Start Date"]).number_format = _DATE_TIME_FORMAT
    ws.cell(row=FIRST_DATA_ROW, column=col["Deadline"]).number_format = _DATE_TIME_FORMAT
    ws.cell(row=FIRST_DATA_ROW, column=col[days_left_name]).number_format = (
        _DAYS_FORMAT if is_manager else _EARLY_LATE_FORMAT
    )
    ws.cell(row=FIRST_DATA_ROW, column=col["Notes"]).number_format = ds.notes_icon_format()
    ws.cell(row=FIRST_DATA_ROW, column=col["Completed At"]).number_format = _DATE_TIME_FORMAT
    if not is_manager:
        ws.cell(row=FIRST_DATA_ROW, column=hidden_col["Time Spent (hrs)"]).number_format = _HOURS_FORMAT

    # §17.11 correction (user-reported): centered alignment for ID, Days
    # Left, Subtask of, Total Subtasks -- same "row 2 sets it, VBA's
    # ListRows.Add inherits it" convention as the number formats above.
    # Manager's Completed Subtasks gets the same treatment, mirroring
    # Tasks' own alignment choice for that column.
    center = Alignment(horizontal="center", vertical="center")
    centered_names = ["ID", "Total Subtasks", "Subtask of", days_left_name]
    if is_manager:
        centered_names.append("Completed Subtasks")
    else:
        centered_names.append("Time Spent")
    for name in centered_names:
        ws.cell(row=FIRST_DATA_ROW, column=all_col[name]).alignment = center

    if is_manager:
        last_col_letter = get_column_letter(hidden_col["RowType"])
    else:
        # Friendly "Xh Ym" Time Spent -- position-relative formula (see
        # this module's own header-block comment for why this is safe
        # despite Archive otherwise being static-value-only), same
        # carry-safe round-to-nearest-minute-first pattern as Tasks' own
        # friendly column.
        ts_ref = f"{get_column_letter(hidden_col['Time Spent (hrs)'])}{FIRST_DATA_ROW}"
        friendly_ts_value_cell = ws.cell(row=FIRST_DATA_ROW, column=col["Time Spent"])
        friendly_ts_value_cell.value = (
            f'=IF({ts_ref}="","",TEXT(INT(ROUND({ts_ref}*60,0)/60),"0")&"h "&TEXT(MOD(ROUND({ts_ref}*60,0),60),"00")&"m")'
        )
        last_col_letter = get_column_letter(hidden_col["Time Spent (hrs)"])

    table = Table(displayName="ArchiveTable", ref=f"A1:{last_col_letter}{FIRST_DATA_ROW}")
    table.tableStyleInfo = ds.modern_table_style()
    ws.add_table(table)

    # --- Conditional formatting: Priority chip (§17.3) + diverging Days
    # Left scale (§17.5) -- CF ranges extend generously below the initial
    # table row since the table (and therefore these rules, which apply
    # to a fixed range, not the ListObject itself) will grow over time.
    # Same rule set for both variants -- manager's "Days Left" freezes the
    # same Deadline-vs-now arithmetic Tasks' own live column uses (plan
    # §6), which happens to carry the identical positive-is-good/
    # negative-is-overdue sign convention the diverging scale already
    # encodes, for both a completed row and a still-open ancestor's
    # Parent-Context snapshot alike -- no new rule needed.
    _CF_LAST_ROW = 1000
    priority_letter = get_column_letter(col["Priority"])
    days_left_letter = get_column_letter(col[days_left_name])
    cf.apply_rules(ws, cf.priority_rules(priority_letter, FIRST_DATA_ROW, _CF_LAST_ROW))
    cf.apply_rules(
        ws, cf.diverging_rules(days_left_letter, FIRST_DATA_ROW, _CF_LAST_ROW, dead_zone=None if is_manager else 0.5)
    )
    # Quiet row banding (§17.10) across the same generous growth range --
    # no per-row hairline borders here (unlike the fixed-size 50-row
    # reference sheets): a direct Border object on every cell of a
    # 1000-row range would mean real per-cell file bloat, unlike a CF
    # rule which is stored once regardless of range size. Banding alone
    # already reads cleanly without it.
    #
    # Manager's PIC column is excluded (same fix as PIC List's own row
    # banding, and Tasks' own depth shading -- user-reported on the
    # latter, same root cause here): a CF fill overrides a cell's own
    # direct fill wherever its condition is true, and banding's condition
    # alternates every other row, so an archived row's PIC color would be
    # invisible on every other row without this.
    banding_pic_col_letter = get_column_letter(col["PIC"]) if is_manager else None
    for banding_range in _visible_ranges_excluding(FIRST_DATA_ROW, _CF_LAST_ROW, last_col_letter, banding_pic_col_letter):
        cf.apply_rules(ws, cf.row_banding_rule(banding_range, FIRST_DATA_ROW))

    if not is_manager:
        # Skipped rows (V4, PLAN.md §22.4) read as "didn't happen": muted
        # italic text across the visible row.
        from openpyxl.formatting.rule import FormulaRule
        from openpyxl.styles import Font

        status_letter = get_column_letter(col["Status"])
        ws.conditional_formatting.add(
            f"A{FIRST_DATA_ROW}:{get_column_letter(len(headers))}{_CF_LAST_ROW}",
            FormulaRule(
                formula=[f'${status_letter}{FIRST_DATA_ROW}="Skipped"'], font=Font(color=ds.INK_MUTED, italic=True)
            ),
        )

    if is_manager:
        # Depth shading + "⤷ " subtask prefix (user-requested after using
        # the manager build: seeing a completed subtask's ancestor chain
        # without any visual tree structure made it hard to scan) --
        # same technique as Tasks' own `_apply_depth_shading`/
        # `_apply_subtask_prefix` (sheets/tasks.py), keyed off Archive's
        # own hidden Level column, which every archived row (Completed or
        # Parent Context alike) already carries a frozen snapshot of.
        # Employee Archive is untouched -- it has no Parent-Context
        # mechanism to visually clarify in the first place, and wasn't
        # part of what was reported.
        level_letter = get_column_letter(hidden_col["Level"])
        task_letter = get_column_letter(col["Task"])
        last_visible_letter = get_column_letter(len(headers))
        pic_letter = get_column_letter(col["PIC"])
        _apply_archive_depth_shading(ws, level_letter, last_visible_letter, _CF_LAST_ROW, exclude_col_letter=pic_letter)
        _apply_archive_subtask_prefix(ws, level_letter, task_letter, _CF_LAST_ROW)

    # --- Sheet protection: locked against direct edits (confirmed with
    # the user, §7.1 correction). UserInterfaceOnly must be re-asserted
    # by Phase 3's Workbook_Open -- see module docstring.
    ws.protection.sheet = True

    # --- Sheet tab color (archive sheet, §17.9) ---
    ws.sheet_properties.tabColor = ds.TAB_ARCHIVE

    return ws


def _visible_ranges_excluding(first_row, last_row, last_col_letter, exclude_col_letter):
    """Splits A{first_row}:{last_col_letter}{last_row} into one or two
    sub-ranges around exclude_col_letter, or returns the whole range
    unsplit if exclude_col_letter is None. Shared by row banding and
    depth shading below -- both need to keep a CF fill (whose condition
    can be true/false per row) off the PIC column specifically, since PIC
    carries its own manager-chosen, persistent direct fill that a CF
    fill would otherwise visually override wherever its own condition
    happens to be true (plan/user feedback -- same root cause fixed for
    PIC List's own row banding and for Tasks' own depth shading).
    """
    if exclude_col_letter is None:
        return [f"A{first_row}:{last_col_letter}{last_row}"]

    from openpyxl.utils import column_index_from_string

    exclude_idx = column_index_from_string(exclude_col_letter)
    last_idx = column_index_from_string(last_col_letter)
    ranges = []
    if exclude_idx > 1:
        ranges.append(f"A{first_row}:{get_column_letter(exclude_idx - 1)}{last_row}")
    if exclude_idx < last_idx:
        ranges.append(f"{get_column_letter(exclude_idx + 1)}{first_row}:{last_col_letter}{last_row}")
    return ranges


def _apply_archive_depth_shading(ws, level_col_letter, last_visible_col_letter, cf_last_row, exclude_col_letter=None):
    """Manager-only (see build()'s own call site). Same technique as
    tasks.py's `_apply_depth_shading` -- progressively lighter row shading
    per level, keyed off the hidden Level column every archived row
    carries a frozen snapshot of. Extends across the same generous
    cf_last_row growth range Archive's other CF rules already use, not
    just the current table size. `exclude_col_letter` (PIC) keeps this
    fill off the PIC column -- see `_visible_ranges_excluding`'s own
    docstring for why.
    """
    from openpyxl.formatting.rule import FormulaRule

    ranges = _visible_ranges_excluding(FIRST_DATA_ROW, cf_last_row, last_visible_col_letter, exclude_col_letter)
    level_cell = f"${level_col_letter}{FIRST_DATA_ROW}"

    shading_steps = [100, 150, 200, 250, 300]  # sequential ramp keys, lightest first
    for depth, step in enumerate(shading_steps, start=1):
        for visible_range in ranges:
            # bgColor, not fgColor -- see conditional_formatting.py's
            # _fill() docstring; this hand-rolled rule bypasses that
            # shared helper so needs the same fix applied directly
            # (matches tasks.py's own). A fresh Rule/PatternFill instance
            # per range, not one object reused across multiple .add()
            # calls (matches tasks.py's own reasoning too).
            fill = PatternFill("solid", bgColor=ds.SEQUENTIAL_BLUE[step])
            rule = FormulaRule(formula=[f"{level_cell}={depth}"], fill=fill)
            ws.conditional_formatting.add(visible_range, rule)


def _apply_archive_subtask_prefix(ws, level_col_letter, task_col_letter, cf_last_row):
    """Manager-only (see build()'s own call site). Same "⤷ " marker as
    tasks.py's `_apply_subtask_prefix` -- a custom TEXT number format
    prefixing the Task cell's own text whenever Level >= 1, so a
    completed subtask (or an ancestor's Parent-Context row one level
    deeper than its own parent) reads as visually nested, matching how
    Tasks itself marks subtask rows.
    """
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.styles.numbers import NumberFormat

    task_range = f"{task_col_letter}{FIRST_DATA_ROW}:{task_col_letter}{cf_last_row}"
    level_cell = f"${level_col_letter}{FIRST_DATA_ROW}"

    rule = Rule(type="expression", formula=[f"{level_cell}>=1"])
    # Custom formats need an explicit numFmtId >= 164 (below that is
    # reserved for Excel's own built-ins) -- 201 matches tasks.py's own
    # choice; distinct sheets don't share a numFmtId namespace so this is
    # safe to reuse verbatim.
    rule.dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=201, formatCode='"⤷ "@'))
    ws.conditional_formatting.add(task_range, rule)
