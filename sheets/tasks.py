"""Tasks -- the most important sheet in the workbook (PLAN.md §6).

This module builds the *structure* only: columns, formatting, conditional
formatting, data validation, and static formulas (Phase 2). The
interactive behaviors -- double-click handlers, the hierarchy engine,
auto-archiving -- are VBA, added in Phase 3 (see vba/modHierarchy.bas,
which this sheet's column layout is built to feed).

Column layout (§6.1), left to right:
  ID, up, down, Start Date, Task, Number of Subtasks, Subtask of,
  Project/Category, Linked Campaign, Priority, Deadline, Status,
  Days Left, Timer, Time Spent, Notes
Hidden columns, grouped together after Notes (not interleaved -- see
hierarchy.py's docstring for why):
  TaskKey, ParentKey, Level, SiblingOrder, TimerRunningSince,
  LinkedCampaignKey, SortKey, CampaignSortHelper, CampaignRank,
  CampaignRankKey

The last three (added when Campaigns was built, §10.2 point 3) let each
campaign's linked-tasks summary block do a single plain INDEX/MATCH per
cell instead of an array-entered (Ctrl+Shift+Enter) SMALL+IF formula --
consistent with how this project avoids CSE formulas everywhere else.
CampaignSortHelper is like SortKey but pushes no-deadline tasks to
"always last" within their campaign group instead of leaving them blank;
CampaignRank is this task's dense rank (1st, 2nd, 3rd soonest...) among
only the tasks sharing its Linked Campaign, via SUMPRODUCT (no CSE
needed); CampaignRankKey combines LinkedCampaignKey & "|" & CampaignRank
into one lookup key so Campaigns can find "the Nth task linked to
campaign X" with a single-condition MATCH.

`Tasks` is a plain pre-formatted range, not a native Excel Table --
unlike Monthly Media Plan's Campaign Table (which is a genuine
auto-expanding ListObject), Tasks' rows get wholesale rewritten in place
by VBA's RenumberHierarchy() on every structural change (§6.2 point 2).
That rewrite strategy was designed and live-tested in Phase 1 against a
plain range; a ListObject's auto-expand semantics would interact with it
in an untested way, so this sheet deliberately stays plain-range-based,
consistent with Monthly Media Plan's Gantt (also plain-range, also
formula-position-sensitive) rather than its Campaign Table.
"""

from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import conditional_formatting as cf
from . import design_system as ds
from .hierarchy import (
    number_of_subtasks_formula,
    write_hidden_hierarchy_columns,
    write_id_reorder_columns,
    write_task_columns,
)

SHEET_NAME = "Tasks"
HEADER_ROW = 1
FIRST_DATA_ROW = 2
LAST_DATA_ROW = 501  # generous pre-formatted allowance, see module docstring

_DATE_TIME_FORMAT = "DD.MM.YYYY HH:MM"
_HOURS_FORMAT = "0.00"
_DAYS_FORMAT = "0"  # whole-day countdown -- see INT() note at days_left_cell below


def _style_header_cell(cell):
    ds.style_header_cell(cell)


def _write_header(ws, col, label):
    cell = ws.cell(row=HEADER_ROW, column=col, value=label)
    _style_header_cell(cell)
    return get_column_letter(col)


def _block_direct_entry(ws, column_letter, message):
    """§6.1: Deadline and Notes are "not directly typed" -- set only via
    the Calendar handoff / frmNotes popup. A custom Data Validation rule
    that's always false blocks interactive typing with a clear error,
    while leaving VBA's programmatic writes unaffected (Data Validation
    only restricts UI-driven input, never code-driven .Value assignment).
    """
    dv = DataValidation(
        type="custom",
        formula1="FALSE",
        showErrorMessage=True,
        errorTitle="Use the picker instead",
        error=message,
    )
    dv.add(f"{column_letter}{FIRST_DATA_ROW}:{column_letter}{LAST_DATA_ROW}")
    ws.add_data_validation(dv)


def build(wb):
    """Create the Tasks sheet in `wb` (an openpyxl Workbook). Returns the
    created worksheet.
    """
    ws = wb.create_sheet(SHEET_NAME)
    ds.apply_modern_chrome(ws)
    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=1)
    ws.row_dimensions[HEADER_ROW].height = ds.HEADER_ROW_HEIGHT

    # --- ID / up / down (columns A-C) ---
    id_cols = write_id_reorder_columns(ws, start_col=1, header_row=HEADER_ROW)

    # --- Start Date (column D) -- sandwiched between the reorder block
    # and Task per §6.1; auto-stamped by VBA (Phase 3), not a formula.
    start_date_col = _write_header(ws, 4, "Start Date")

    # --- Task / Number of Subtasks / Subtask of (columns E-G) ---
    task_cols = write_task_columns(ws, start_col=5, header_row=HEADER_ROW)

    # --- Remaining visible columns (H onward) ---
    project_col = _write_header(ws, 8, "Project / Category")
    linked_campaign_col = _write_header(ws, 9, "Linked Campaign")
    priority_col = _write_header(ws, 10, "Priority")
    deadline_col = _write_header(ws, 11, "Deadline")
    status_col = _write_header(ws, 12, "Status")
    days_left_col = _write_header(ws, 13, "Days Left")
    timer_col = _write_header(ws, 14, "Timer")
    time_spent_col = _write_header(ws, 15, "Time Spent (hrs)")
    notes_col = _write_header(ws, 16, "Notes")

    # --- Hidden columns, grouped together starting at column Q ---
    hidden_cols = write_hidden_hierarchy_columns(
        ws, start_col=17, key_col_name="TaskKey", parent_key_col_name="ParentKey", header_row=HEADER_ROW
    )
    timer_running_since_col = _write_header(ws, 21, "TimerRunningSince")
    linked_campaign_key_col = _write_header(ws, 22, "LinkedCampaignKey")
    sort_key_col = _write_header(ws, 23, "SortKey")
    campaign_sort_helper_col = _write_header(ws, 24, "CampaignSortHelper")
    campaign_rank_col = _write_header(ws, 25, "CampaignRank")
    campaign_rank_key_col = _write_header(ws, 26, "CampaignRankKey")

    # --- Friendly Time Spent display (§19 Phase 5 correction) -- appended
    # after the existing self-healing helper block rather than inserted
    # next to Time Spent (hrs) itself, to avoid renumbering every column
    # after it across modTasks.bas/modRecurringTasks.bas/campaigns.py's
    # cross-sheet references (confirmed with the user: lower risk than
    # true proximity to Time Spent (hrs)). A plain number FORMAT can't
    # produce this -- Excel's "[h]:mm:ss" duration tokens always assume
    # the stored value is a fraction of a DAY, but Time Spent (hrs) is
    # stored in HOURS, so this needs an actual formula, not a format
    # string. Position-relative (references only its own row's Time Spent
    # (hrs) cell), so it's just as self-healing as SortKey/etc above --
    # no VBA changes needed, RenumberHierarchy already excludes anything
    # past LastCol=22.
    friendly_time_spent_col = _write_header(ws, 27, "Time Spent")

    for letter in (
        timer_running_since_col,
        linked_campaign_key_col,
        sort_key_col,
        campaign_sort_helper_col,
        campaign_rank_col,
        campaign_rank_key_col,
        time_spent_col,
    ):
        ws.column_dimensions[letter].hidden = True

    # --- Column widths (visible columns only) ---
    widths = {
        id_cols["id"]: 10,
        id_cols["up"]: 4,
        id_cols["down"]: 4,
        # 19, not 15 (§17.10 bugfix) -- "DD.MM.YYYY HH:MM" is 16
        # characters, which the 115%-zoom modernization pass (this same
        # Phase 6 pass) pushed past a width-15 column's fit, rendering as
        # "#####" -- user-reported live. Column width is stored in
        # character units and doesn't scale with zoom, so a column
        # already razor-thin for its content at 100% overflows once
        # zoom makes every character render larger on screen.
        start_date_col: 19,
        task_cols["task"]: 32,
        task_cols["num_subtasks"]: 10,
        task_cols["subtask_of"]: 10,
        project_col: 18,
        linked_campaign_col: 18,
        priority_col: 10,
        deadline_col: 19,
        status_col: 13,
        days_left_col: 10,
        timer_col: 10,
        time_spent_col: 10,
        notes_col: 8,
        friendly_time_spent_col: 12,
    }
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width

    # --- Number formats + static formulas across the data range ---
    for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
        ws.cell(row=row, column=4).number_format = _DATE_TIME_FORMAT  # Start Date
        ws.cell(row=row, column=11).number_format = _DATE_TIME_FORMAT  # Deadline
        ws.cell(row=row, column=15).number_format = _HOURS_FORMAT  # Time Spent
        ws.cell(row=row, column=16).number_format = ds.notes_icon_format()  # Notes

        num_sub_cell = ws.cell(row=row, column=6)
        num_sub_cell.value = number_of_subtasks_formula(
            id_cols["id"], row, FIRST_DATA_ROW, LAST_DATA_ROW
        )

        days_left_cell = ws.cell(row=row, column=13)
        # §20 bugfix: Deadline always carries a time-of-day component (set
        # via Calendar's date+time picker, or TimeSerial(18,0,0) on a
        # Recurring Tasks generated instance -- never midnight), so a raw
        # Deadline-TODAY() is a decimal, and the "0" number format then
        # ROUNDS that decimal for display -- a deadline of today 18:00 read
        # back as 0.75, which the "0" format rounds UP to display "1", even
        # though the task is due today (0 days left), not tomorrow. INT()
        # strips the time-of-day off Deadline first, so the subtraction is a
        # true whole-calendar-day countdown and the stored value is already
        # an integer -- no rounding ambiguity for CF/sorting/Dashboard, all
        # of which read this same cell.
        days_left_cell.value = f'=IF({deadline_col}{row}="","",INT({deadline_col}{row})-TODAY())'
        days_left_cell.number_format = _DAYS_FORMAT

        # Friendly "1h 12m" breakdown of Time Spent (hrs) -- round to the
        # nearest whole minute FIRST, then split into hours/minutes, so a
        # value like 1.999h correctly carries into "2h 00m" instead of
        # rounding each part separately and landing on the nonsensical
        # "1h 60m".
        friendly_time_spent_cell = ws.cell(row=row, column=27)
        ts = f"{time_spent_col}{row}"
        friendly_time_spent_cell.value = (
            f'=IF({ts}="","",TEXT(INT(ROUND({ts}*60,0)/60),"0")&"h '
            f'"&TEXT(MOD(ROUND({ts}*60,0),60),"00")&"m")'
        )

        # §4.B (corrected): SortKey = Deadline + TaskKey/1e6, a unique,
        # sortable tie-breaker for Dashboard's top-N extraction.
        sort_key_cell = ws.cell(row=row, column=23)
        sort_key_cell.value = (
            f'=IF({deadline_col}{row}="","",{deadline_col}{row}+{hidden_cols["key"]}{row}/1000000)'
        )

        # §10.2 point 3: per-campaign top-N lookup helpers (see module
        # docstring). CampaignSortHelper pushes no-deadline tasks to
        # "always last" within their group (9,999,999 is far beyond any
        # real date serial, safe as an "infinity" stand-in).
        key = hidden_cols["key"]
        lck = f"{linked_campaign_key_col}{row}"
        csh_range = f"${campaign_sort_helper_col}${FIRST_DATA_ROW}:${campaign_sort_helper_col}${LAST_DATA_ROW}"
        lck_range = f"${linked_campaign_key_col}${FIRST_DATA_ROW}:${linked_campaign_key_col}${LAST_DATA_ROW}"

        ws.cell(row=row, column=24).value = (
            f'=IF({lck}="","",IF({deadline_col}{row}="",9999999+{key}{row}/1000000,'
            f'{deadline_col}{row}+{key}{row}/1000000))'
        )
        ws.cell(row=row, column=25).value = (
            f'=IF({lck}="","",SUMPRODUCT(({lck_range}={lck})*'
            f'({csh_range}<={campaign_sort_helper_col}{row})))'
        )
        ws.cell(row=row, column=26).value = f'=IF({lck}="","",{lck}&"|"&{campaign_rank_col}{row})'

    # --- Data Validation dropdowns ---
    priority_dv = DataValidation(type="list", formula1='"Low,Medium,High"', allow_blank=True)
    priority_dv.add(f"{priority_col}{FIRST_DATA_ROW}:{priority_col}{LAST_DATA_ROW}")
    ws.add_data_validation(priority_dv)

    status_dv = DataValidation(
        type="list", formula1='"Not Started,In Progress,Completed,Delete"', allow_blank=True
    )
    status_dv.add(f"{status_col}{FIRST_DATA_ROW}:{status_col}{LAST_DATA_ROW}")
    ws.add_data_validation(status_dv)

    # Linked Campaign dropdown, wired now that Campaigns exists (2E). Its
    # source is a named range ("CampaignNameList") rather than a direct
    # reference to Campaigns' Campaign Name column, because that column's
    # 20 slots aren't contiguous (each reserves 8 hidden summary rows
    # beneath it, §10.2 point 4) and Data Validation lists can't reference
    # a non-contiguous range. campaigns.py defines CampaignNameList as a
    # contiguous 20-cell mirror of the 20 (non-contiguous) slot cells.
    linked_campaign_dv = DataValidation(type="list", formula1="=CampaignNameList", allow_blank=True)
    linked_campaign_dv.add(f"{linked_campaign_col}{FIRST_DATA_ROW}:{linked_campaign_col}{LAST_DATA_ROW}")
    ws.add_data_validation(linked_campaign_dv)

    # Deadline and Notes are "not directly typed" (§6.1, §6.3) -- set only
    # via the Calendar handoff / frmNotes popup.
    _block_direct_entry(
        ws, deadline_col, "Double-click this cell to pick a date from the Calendar instead."
    )
    _block_direct_entry(
        ws, notes_col, "Double-click this cell to open the Notes popup instead."
    )

    # --- Conditional formatting ---
    cf.apply_rules(ws, cf.priority_rules(priority_col, FIRST_DATA_ROW, LAST_DATA_ROW))
    cf.apply_rules(ws, cf.urgency_rules(days_left_col, FIRST_DATA_ROW, LAST_DATA_ROW))
    cf.apply_rules(
        ws,
        cf.workflow_status_rules(
            status_col, FIRST_DATA_ROW, LAST_DATA_ROW, not_started_label="Not Started", in_motion_label="In Progress"
        ),
    )
    _apply_depth_shading(ws, hidden_cols["level"])
    _apply_subtask_prefix(ws, hidden_cols["level"], task_cols["task"])

    # Row structure (§17.12 correction, user-reported: "the information
    # is floating and there is no clear row structure") -- a hairline
    # bottom border across every visible column, every data row.
    ds.apply_table_hairline_borders(ws, 1, 27, FIRST_DATA_ROW, LAST_DATA_ROW)  # through the friendly Time Spent column

    # Button-shaped up/down/Timer cells (§17.11) are deliberately NOT
    # applied here at build time -- §17.13's correction (user-reported:
    # a build-time version styled all 500 pre-formatted rows identically,
    # so the button borders visibly continued downward through hundreds
    # of rows with no task in them). Applied dynamically instead, per row,
    # the moment a row actually becomes a task -- see modTasks.bas's
    # RefreshButtonStyling, called after every RenumberHierarchy site.

    # Centered ID, Total Subtasks, Subtask of, Days Left, Time Spent
    # (§17.11/§17.14, user-reported) -- baked in across the full data
    # range at build time, same reasoning throughout: VBA
    # (RenumberHierarchy) only ever rewrites a cell's *value*, never its
    # formatting, so a fixed alignment set here at a fixed row/column
    # position survives any amount of hierarchy reshuffling.
    center = Alignment(horizontal="center", vertical="center")
    for col_letter in (
        id_cols["id"],
        task_cols["num_subtasks"],
        task_cols["subtask_of"],
        days_left_col,
        friendly_time_spent_col,
    ):
        for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
            ws[f"{col_letter}{row}"].alignment = center

    # Sequential blue ramp on Number of Subtasks per §17.5 (PLAN.md §19
    # Phase 6 -- confirmed with the user this was missing entirely,
    # simple magnitude ramp chosen over the original §6.4 draft's
    # done-vs-open proportional idea). A 0-subtask row shows "" (not 0,
    # see the formula this column already uses), so the ISNUMBER guard
    # inside sequential_rules() already keeps it unfilled -- no explicit
    # zero bucket needed. Same light 100-300 range as depth shading.
    cf.apply_rules(
        ws,
        cf.sequential_rules(
            task_cols["num_subtasks"],
            FIRST_DATA_ROW,
            LAST_DATA_ROW,
            buckets=[(1, 100), (3, 150), (6, 200), (11, 250), (21, 300)],
        ),
    )

    # --- Sheet tab color (core daily-use sheet, §17.9) ---
    ws.sheet_properties.tabColor = ds.TAB_CORE

    return ws


def _apply_depth_shading(ws, level_col_letter):
    """§6.4: "progressively lighter row shading per level, so nesting is
    readable without visual clutter." Applied across the full visible
    column range (A:P), keyed off the hidden Level column -- a secondary,
    subtle reinforcement of depth alongside the Task cell's own
    IndentLevel (which VBA sets, since indentation isn't formula-drivable).
    Uses the lightest steps of the sequential blue ramp (§17.5); Level 0
    (top-level tasks) gets no fill.
    """
    from openpyxl.formatting.rule import FormulaRule

    visible_range = f"A{FIRST_DATA_ROW}:P{LAST_DATA_ROW}"
    level_cell = f"${level_col_letter}{FIRST_DATA_ROW}"

    shading_steps = [100, 150, 200, 250, 300]  # sequential ramp keys, lightest first
    for depth, step in enumerate(shading_steps, start=1):
        # bgColor, not fgColor -- see conditional_formatting.py's _fill()
        # docstring; this hand-rolled rule bypasses that shared helper so
        # needs the same fix applied directly.
        fill = PatternFill("solid", bgColor=ds.SEQUENTIAL_BLUE[step])
        rule = FormulaRule(formula=[f"{level_cell}={depth}"], fill=fill)
        ws.conditional_formatting.add(visible_range, rule)


def _apply_subtask_prefix(ws, level_col_letter, task_col_letter):
    """User-requested addition (PLAN.md §19 Phase 5): a "⤷ " prefix on
    every subtask row (Level >= 1), so a row's hierarchy status is
    visible at a glance without relying on indentation alone. Purely
    conditional-formatting-driven -- a custom TEXT number format,
    '"⤷ "@', prefixing the cell's own text -- NOT baked into the Task
    cell's actual stored value, so no VBA is needed at all: it appears or
    disappears automatically as RenumberHierarchy's live Level value
    changes when a task is reparented, promoted, or demoted, exactly like
    the depth-shading fill above it (same hidden Level column, same
    "recalculates live" property). Same marker at every depth (confirmed
    with the user, chosen from a few prefix-style previews) rather than
    repeating it per level.
    """
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.styles.numbers import NumberFormat

    task_range = f"{task_col_letter}{FIRST_DATA_ROW}:{task_col_letter}{LAST_DATA_ROW}"
    level_cell = f"${level_col_letter}{FIRST_DATA_ROW}"

    rule = Rule(type="expression", formula=[f"{level_cell}>=1"])
    # Custom formats need an explicit numFmtId >= 164 (below that is
    # reserved for Excel's own built-ins) -- 201 is arbitrary but must be
    # present, openpyxl errors on numFmtId=None.
    rule.dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=201, formatCode='"⤷ "@'))
    ws.conditional_formatting.add(task_range, rule)
