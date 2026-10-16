"""Tasks -- the most important sheet in the workbook (PLAN.md §6).

This module builds the *structure* only: columns, formatting, conditional
formatting, data validation, and static formulas (Phase 2). The
interactive behaviors -- double-click handlers, the hierarchy engine,
auto-archiving -- are VBA, added in Phase 3 (see vba/modHierarchy.bas,
which this sheet's column layout is built to feed).

`build(wb, variant="employee")` supports two layouts sharing this one
module rather than a forked file, per the manager-variant design (PLAN.md,
"Manager-variant workbook"): the `variant="employee"` path is untouched
byte-for-byte from before this parameter existed (verified: build output
identical for `variant="employee"` and the old no-arg `build(wb)`).

Employee column layout (§6.1), left to right:
  ID, up, down, Start Date, Task, Number of Subtasks, Subtask of,
  Project/Category, Linked Campaign, Priority, Deadline, Status,
  Days Left, Timer, Time Spent, Notes
Hidden columns, grouped together after Notes (not interleaved -- see
hierarchy.py's docstring for why):
  TaskKey, ParentKey, Level, SiblingOrder, TimerRunningSince,
  LinkedCampaignKey, SortKey, CampaignSortHelper, CampaignRank,
  CampaignRankKey
Followed by one more visible column, Time Spent (friendly "1h 12m").

Manager column layout: identical except PIC is inserted right after Task,
Completed Subtasks is inserted right after Number of Subtasks, and Timer /
Time Spent / TimerRunningSince / the friendly Time Spent column are all
omitted entirely (the manager variant tracks neither a timer nor hours).
PIC is a plain dropdown-sourced value with no hidden key (unlike Linked
Campaign, a person's name has no rename-tracking need) -- so it needs no
extra hidden column of its own. Completed Subtasks is a plain VBA-written
running counter (like Time Spent is on the employee sheet), not a live
formula -- completed subtasks are archived out of Tasks immediately
(archive-on-complete), so unlike Number of Subtasks there's nothing left
in Tasks for a COUNTIF-style formula to count. By coincidence of column
counts (PIC + Completed Subtasks added, Timer + Time Spent removed), both
variants' visible block spans the same A:P range and their hidden block
starts at the same column Q -- confirmed deliberately, not incidental,
since several formulas/CF ranges below assume a shared A:P visible span.

The last three of the shared hidden-column tail (added when Campaigns was
built, §10.2 point 3) let each campaign's linked-tasks summary block do a
single plain INDEX/MATCH per cell instead of an array-entered (Ctrl+Shift+
Enter) SMALL+IF formula -- consistent with how this project avoids CSE
formulas everywhere else. CampaignSortHelper is like SortKey but pushes
no-deadline tasks to "always last" within their campaign group instead of
leaving them blank; CampaignRank is this task's dense rank (1st, 2nd,
3rd soonest...) among only the tasks sharing its Linked Campaign, via
SUMPRODUCT (no CSE needed); CampaignRankKey combines LinkedCampaignKey &
"|" & CampaignRank into one lookup key so Campaigns can find "the Nth task
linked to campaign X" with a single-condition MATCH. Both variants keep
Campaigns/Linked Campaign, so both need this same helper chain.

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

from openpyxl.styles import Alignment, Font, PatternFill
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
LAST_DATA_ROW = 1500  # generous pre-formatted allowance, see module docstring (raised from 500, user ran out)

_DATE_TIME_FORMAT = "DD.MM.YYYY HH:MM"
_DATE_FORMAT = "DD.MM.YYYY"
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


def build(wb, variant="employee"):
    """Create the Tasks sheet in `wb` (an openpyxl Workbook). `variant` is
    "employee" (default, unchanged from before this parameter existed) or
    "manager" (PIC + Completed Subtasks columns, no Timer/Time Spent).
    Returns the created worksheet.
    """
    assert variant in ("employee", "manager"), f"unknown variant: {variant!r}"
    is_manager = variant == "manager"

    ws = wb.create_sheet(SHEET_NAME)
    ds.apply_modern_chrome(ws)
    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=1)
    ws.row_dimensions[HEADER_ROW].height = ds.HEADER_ROW_HEIGHT

    # --- ID / up / down (columns A-C) -- identical for both variants ---
    id_cols = write_id_reorder_columns(ws, start_col=1, header_row=HEADER_ROW)

    # --- Start Date (column D) -- sandwiched between the reorder block
    # and Task per §6.1; auto-stamped by VBA (Phase 3), not a formula.
    start_date_col = _write_header(ws, 4, "Start Date")

    # --- Task onward: the two variants' visible layouts diverge here, but
    # both span exactly columns E through P (16 visible columns total,
    # A-P), so the hidden block below starts at the same column Q either
    # way -- see module docstring.
    if is_manager:
        task_col = _write_header(ws, 5, "Task")
        pic_col = _write_header(ws, 6, "PIC")
        num_subtasks_col = _write_header(ws, 7, "Total Subtasks")
        completed_subtasks_col = _write_header(ws, 8, "Completed Subtasks")
        subtask_of_col = _write_header(ws, 9, "Subtask of")
        task_cols = {
            "task": task_col,
            "num_subtasks": num_subtasks_col,
            "subtask_of": subtask_of_col,
        }
        project_col = _write_header(ws, 10, "Project / Category")
        linked_campaign_col = _write_header(ws, 11, "Linked Campaign")
        priority_col = _write_header(ws, 12, "Priority")
        # "Deadline ↓" (user-requested "sort by deadline" button, manager
        # only for now): double-clicking this header triggers
        # modTasksManager.bas's HandleSortByDeadlineClick, reordering
        # tasks by ascending deadline WITHIN each existing PIC-cluster/
        # parent group (soonest first, blank last) rather than blowing
        # away PIC grouping -- the ↓ marks it as clickable/sortable, same
        # convention as Up/Down's own arrow glyphs elsewhere on this
        # sheet.
        deadline_col = _write_header(ws, 13, "Deadline")
        status_col = _write_header(ws, 14, "Status")
        days_left_col = _write_header(ws, 15, "Days Left")
        notes_col = _write_header(ws, 16, "Notes")
        timer_col = None
        time_spent_col = None
    else:
        task_cols = write_task_columns(ws, start_col=5, header_row=HEADER_ROW)
        pic_col = None
        completed_subtasks_col = None
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
    if is_manager:
        # No TimerRunningSince -- manager tracks no timer at all.
        linked_campaign_key_col = _write_header(ws, 21, "LinkedCampaignKey")
        sort_key_col = _write_header(ws, 22, "SortKey")
        campaign_sort_helper_col = _write_header(ws, 23, "CampaignSortHelper")
        campaign_rank_col = _write_header(ws, 24, "CampaignRank")
        campaign_rank_key_col = _write_header(ws, 25, "CampaignRankKey")
        timer_running_since_col = None
        friendly_time_spent_col = None
        last_col = 25
    else:
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
        last_col = 27
        # Hidden RecurringSourceKey (V4, PLAN.md §22.4): the Recurring
        # Tasks template root's RecurringKey, written on a GENERATED root
        # row only, so "If Missed = Skip" can find a template's still-open
        # instances. A moved value column, so it sits after the formula
        # helpers and modTasks' TasksColMap marks 23-27 as KeepFrom/KeepTo
        # (never rewritten) while still carrying column 28 with its row.
        recurring_source_key_col = _write_header(ws, 28, "RecurringSourceKey")
        ws.column_dimensions[recurring_source_key_col].hidden = True

        # Subtask progress (V4, PLAN.md §22.11). DoneSubtasks (hidden) is a
        # VBA-maintained running count of completed/skipped descendants --
        # a finished subtask leaves Tasks for Archive, so nothing else on
        # this sheet remembers it. Inside TasksColMap's LastCol so it moves
        # with its row. Progress (visible) is a per-row display formula
        # past LastCol, like Time Spent. Both sit at the far right rather
        # than next to Total Subtasks: inserting a column mid-sheet would
        # shift every hard-coded Tasks column in the VBA, Dashboard and
        # Campaigns.
        done_subtasks_col = _write_header(ws, 29, "DoneSubtasks")
        ws.column_dimensions[done_subtasks_col].hidden = True
        progress_col = _write_header(ws, 30, "Progress")
        last_col = 30

    # The 4 hierarchy columns (key/parent_key/level/sibling_order) are
    # already marked hidden by write_hidden_hierarchy_columns itself --
    # this list is everything else that needs it.
    hidden_letters = [
        linked_campaign_key_col,
        sort_key_col,
        campaign_sort_helper_col,
        campaign_rank_col,
        campaign_rank_key_col,
    ]
    if not is_manager:
        hidden_letters.append(timer_running_since_col)
        hidden_letters.append(time_spent_col)
    for letter in hidden_letters:
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
        notes_col: 8,
    }
    if is_manager:
        widths[pic_col] = 16
        widths[completed_subtasks_col] = 12
    else:
        widths[timer_col] = 10
        widths[time_spent_col] = 10
        widths[friendly_time_spent_col] = 12
        widths[progress_col] = 16
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width

    # --- Number formats + static formulas across the data range ---
    for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
        # Start Date: date only (V4, PLAN.md §22.8) -- the creation time
        # of day was noise on every row. The value keeps its time.
        ws.cell(row=row, column=4).number_format = _DATE_TIME_FORMAT if is_manager else _DATE_FORMAT  # Start Date
        ws.cell(row=row, column=column_index(deadline_col)).number_format = _DATE_TIME_FORMAT  # Deadline
        ws.cell(row=row, column=column_index(notes_col)).number_format = ds.notes_icon_format()  # Notes
        if not is_manager:
            ws.cell(row=row, column=column_index(time_spent_col)).number_format = _HOURS_FORMAT  # Time Spent

        num_sub_cell = ws.cell(row=row, column=column_index(task_cols["num_subtasks"]))
        num_sub_cell.value = number_of_subtasks_formula(
            id_cols["id"], row, FIRST_DATA_ROW, LAST_DATA_ROW
        )
        # Completed Subtasks (manager only) is a plain VBA-written running
        # counter, not a formula -- see module docstring. Left blank at
        # build time (matches Total Subtasks' own "blank not zero"
        # convention for an unused/never-touched row).

        days_left_cell = ws.cell(row=row, column=column_index(days_left_col))
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

        if not is_manager:
            # Friendly "1h 12m" breakdown of Time Spent (hrs) -- round to the
            # nearest whole minute FIRST, then split into hours/minutes, so a
            # value like 1.999h correctly carries into "2h 00m" instead of
            # rounding each part separately and landing on the nonsensical
            # "1h 60m".
            friendly_time_spent_cell = ws.cell(row=row, column=column_index(friendly_time_spent_col))
            ts = f"{time_spent_col}{row}"
            friendly_time_spent_cell.value = (
                f'=IF({ts}="","",TEXT(INT(ROUND({ts}*60,0)/60),"0")&"h '
                f'"&TEXT(MOD(ROUND({ts}*60,0),60),"00")&"m")'
            )

        if not is_manager:
            progress_cell = ws.cell(row=row, column=column_index(progress_col))
            progress_cell.value = progress_formula(
                hidden_cols["key"], task_cols["num_subtasks"], done_subtasks_col, row
            )
            progress_cell.alignment = Alignment(horizontal="center", vertical="center")

        # §4.B (corrected): SortKey = Deadline + TaskKey/1e6, a unique,
        # sortable tie-breaker for Dashboard's top-N extraction.
        sort_key_cell = ws.cell(row=row, column=column_index(sort_key_col))
        sort_key_cell.value = (
            f'=IF({deadline_col}{row}="","",{deadline_col}{row}+{hidden_cols["key"]}{row}/1000000)'
        )

        # §10.2 point 3: per-campaign top-N lookup helpers (see module
        # docstring). CampaignSortHelper pushes no-deadline tasks to
        # "always last" within their group (9,999,999 is far beyond any
        # real date serial, safe as an "infinity" stand-in). Both variants
        # keep Campaigns/Linked Campaign, so both need this chain.
        key = hidden_cols["key"]
        lck = f"{linked_campaign_key_col}{row}"
        csh_range = f"${campaign_sort_helper_col}${FIRST_DATA_ROW}:${campaign_sort_helper_col}${LAST_DATA_ROW}"
        lck_range = f"${linked_campaign_key_col}${FIRST_DATA_ROW}:${linked_campaign_key_col}${LAST_DATA_ROW}"

        ws.cell(row=row, column=column_index(campaign_sort_helper_col)).value = (
            f'=IF({lck}="","",IF({deadline_col}{row}="",9999999+{key}{row}/1000000,'
            f'{deadline_col}{row}+{key}{row}/1000000))'
        )
        ws.cell(row=row, column=column_index(campaign_rank_col)).value = (
            f'=IF({lck}="","",SUMPRODUCT(({lck_range}={lck})*'
            f'({csh_range}<={campaign_sort_helper_col}{row})))'
        )
        ws.cell(row=row, column=column_index(campaign_rank_key_col)).value = (
            f'=IF({lck}="","",{lck}&"|"&{campaign_rank_col}{row})'
        )

    # --- Data Validation dropdowns ---
    priority_dv = DataValidation(type="list", formula1='"Low,Medium,High"', allow_blank=True)
    priority_dv.add(f"{priority_col}{FIRST_DATA_ROW}:{priority_col}{LAST_DATA_ROW}")
    ws.add_data_validation(priority_dv)

    # Skipped / Rescheduled are employee-only (V4, PLAN.md §22.4).
    status_values = (
        "Not Started,In Progress,Completed,Delete"
        if is_manager
        else "Not Started,In Progress,Completed,Skipped,Rescheduled,Delete"
    )
    status_dv = DataValidation(type="list", formula1=f'"{status_values}"', allow_blank=True)
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

    if is_manager:
        # PIC dropdown -- sourced from the maintained, growable "PIC List"
        # sheet (pic_list.py) via the same dynamic-range technique as
        # CampaignNameList, for the same documented reason (a plain
        # fixed-size range can leave a trailing selectable blank row).
        # Unlike Linked Campaign, no hidden key column is needed -- a
        # person's name has no rename-tracking concern.
        pic_dv = DataValidation(type="list", formula1="=PICList", allow_blank=True)
        pic_dv.add(f"{pic_col}{FIRST_DATA_ROW}:{pic_col}{LAST_DATA_ROW}")
        ws.add_data_validation(pic_dv)

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
    if not is_manager:
        cf.apply_rules(
            ws,
            cf.categorical_rules(
                status_col,
                FIRST_DATA_ROW,
                LAST_DATA_ROW,
                [("Skipped", ds.WORKFLOW_SKIPPED), ("Rescheduled", ds.WORKFLOW_RESCHEDULED)],
            ),
        )
    if not is_manager:
        # After the chips, so Priority/Status/Days Left keep their fills.
        _apply_v4_task_cues(ws, deadline_col, linked_campaign_col)
        _apply_progress_complete(ws, progress_col, task_cols["num_subtasks"], done_subtasks_col)
    _apply_depth_shading(ws, hidden_cols["level"], exclude_col_letter=pic_col if is_manager else None)
    _apply_subtask_prefix(ws, hidden_cols["level"], task_cols["task"])

    # Row structure (§17.12 correction, user-reported: "the information
    # is floating and there is no clear row structure") -- a hairline
    # bottom border across every visible column, every data row. Employee
    # extends through the friendly Time Spent column (27); manager has no
    # equivalent trailing visible column, so it stops at Notes (16) -- the
    # columns in between (Q onward) are hidden either way, so this is a
    # cosmetic-only difference.
    ds.apply_table_hairline_borders(ws, 1, last_col if not is_manager else 16, FIRST_DATA_ROW, LAST_DATA_ROW)

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
    # position survives any amount of hierarchy reshuffling. Manager's
    # Completed Subtasks gets the same treatment as Total Subtasks, for
    # the same reason.
    center = Alignment(horizontal="center", vertical="center")
    centered_cols = [
        id_cols["id"],
        task_cols["num_subtasks"],
        task_cols["subtask_of"],
        days_left_col,
    ]
    if is_manager:
        centered_cols.append(completed_subtasks_col)
    else:
        centered_cols.append(friendly_time_spent_col)
    for col_letter in centered_cols:
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

    # --- Row extender banner + row-allowance bookkeeping (user-requested,
    # "keep this tracker usable for years, not just until the rows run
    # out") -- a permanent, re-runnable "add 500 more rows" button rather
    # than a one-off manual row-count bump. TasksRowAllowance is a
    # workbook-scoped Defined Name holding a bare constant (not a cell/
    # range reference) -- the single source of truth for "how many rows
    # are currently allotted," read by modTasks.bas's extender AND
    # modCampaigns.bas's ResolveMiniListTaskNotesCell (converted from a
    # hardcoded Const to a Function reading this Name) so neither ever
    # needs its own bound patched by hand again. The banner cell always
    # sits exactly one row below the current allowance; the extender
    # macro moves it down each time it runs. See PLAN.md's row-extender
    # design note for the full list of formulas this bound touches.
    from openpyxl.workbook.defined_name import DefinedName

    wb.defined_names["TasksRowAllowance"] = DefinedName("TasksRowAllowance", attr_text=str(LAST_DATA_ROW))

    banner_row = LAST_DATA_ROW + 1
    ws.merge_cells(f"A{banner_row}:F{banner_row}")
    banner = ws.cell(row=banner_row, column=1, value="+ Add 500 More Task Rows")
    banner.font = Font(name=ds.FONT_NAME, size=ds.FONT_SIZE_BODY, bold=True, color=ds.INK_PRIMARY)
    banner.fill = PatternFill("solid", fgColor=ds.CARD_SURFACE)
    ds.apply_box_outline(ws, f"A{banner_row}:F{banner_row}", color=ds.INK_MUTED)
    banner.alignment = Alignment(horizontal="center", vertical="center")

    # --- Sheet tab color (core daily-use sheet, §17.9) ---
    ws.sheet_properties.tabColor = ds.TAB_CORE

    return ws


def column_index(letter):
    """Small local helper -- the manager/employee branches above compute
    several column *letters* via variables rather than fixed literals, but
    the per-row formula loop needs numeric column indices for
    `ws.cell(row=..., column=...)`. openpyxl's own column_index_from_string
    does exactly this; wrapped under a shorter local name since it's called
    many times in the loop above.
    """
    from openpyxl.utils import column_index_from_string

    return column_index_from_string(letter)


def _apply_depth_shading(ws, level_col_letter, exclude_col_letter=None):
    """§6.4: "progressively lighter row shading per level, so nesting is
    readable without visual clutter." Applied across the full visible
    column range (A:P -- both variants' visible block spans exactly this
    range, see module docstring), keyed off the hidden Level column -- a
    secondary, subtle reinforcement of depth alongside the Task cell's own
    IndentLevel (which VBA sets, since indentation isn't formula-drivable).
    Uses the lightest steps of the sequential blue ramp (§17.5); Level 0
    (top-level tasks) gets no fill.

    `exclude_col_letter` (manager only -- PIC) splits the range in two
    around that one column instead of applying uniformly, user-reported:
    a PIC's manually-set cell color displayed fine at Level 0 but got
    visually overridden the moment a task became a subtask. Same root
    cause as PIC List's own row-banding fix -- a conditional-format fill
    overrides a cell's own direct fill wherever the CF condition is true,
    and this rule's condition (Level >= 1) is exactly "became a subtask."
    A PIC color is manager-chosen, persistent data (like the PIC List
    swatch it's read from), so depth shading -- like row banding there --
    can't be allowed to paint over it.
    """
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.utils import column_index_from_string

    level_cell = f"${level_col_letter}{FIRST_DATA_ROW}"

    if exclude_col_letter is None:
        ranges = [f"A{FIRST_DATA_ROW}:P{LAST_DATA_ROW}"]
    else:
        exclude_idx = column_index_from_string(exclude_col_letter)
        ranges = []
        if exclude_idx > 1:
            before_letter = get_column_letter(exclude_idx - 1)
            ranges.append(f"A{FIRST_DATA_ROW}:{before_letter}{LAST_DATA_ROW}")
        if exclude_idx < 16:
            after_letter = get_column_letter(exclude_idx + 1)
            ranges.append(f"{after_letter}{FIRST_DATA_ROW}:P{LAST_DATA_ROW}")

    shading_steps = [100, 150, 200, 250, 300]  # sequential ramp keys, lightest first
    for depth, step in enumerate(shading_steps, start=1):
        for rng in ranges:
            # bgColor, not fgColor -- see conditional_formatting.py's
            # _fill() docstring; this hand-rolled rule bypasses that
            # shared helper so needs the same fix applied directly. A
            # fresh Rule/PatternFill instance per range, not one object
            # reused across multiple .add() calls, to avoid any
            # openpyxl internal-state assumption about a rule only ever
            # being added once.
            fill = PatternFill("solid", bgColor=ds.SEQUENTIAL_BLUE[step])
            rule = FormulaRule(formula=[f"{level_cell}={depth}"], fill=fill)
            ws.conditional_formatting.add(rng, rule)


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


def _apply_v4_task_cues(ws, deadline_col, linked_campaign_col):
    """Employee-only visual cues (V4, PLAN.md §22.8):
    - the ID cell turns red (white text) when the task is overdue;
    - Deadline reads "Today 18:00" / "Tomorrow 10:00" / "Yesterday 09:00"
      for the three nearest days, via conditional NUMBER FORMATS -- the
      stored value is untouched, so sorting, the Calendar picker and every
      formula reading Deadline keep working on the real date;
    - Linked Campaign's placeholder "None" is shown faint instead of
      competing with real campaign names on almost every row.
    """
    from openpyxl.formatting.rule import FormulaRule, Rule
    from openpyxl.styles import Font
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.styles.numbers import NumberFormat

    k = f"${deadline_col}{FIRST_DATA_ROW}"
    has_deadline = f"ISNUMBER({k})"

    ws.conditional_formatting.add(
        f"A{FIRST_DATA_ROW}:A{LAST_DATA_ROW}",
        FormulaRule(
            formula=[f"AND({has_deadline},INT({k})<TODAY())"],
            fill=PatternFill("solid", bgColor=ds.STATUS_CRITICAL),
            font=Font(color=ds.BRAND_ACCENT_TEXT, bold=True),
        ),
    )
    # No whole-row "due today" tint (removed after user testing, PLAN.md
    # §22.12): it outranked the subtask depth shading, so a parent and its
    # subtasks due the same day all turned the same colour and the
    # hierarchy disappeared. "Today hh:mm" (below) + Days Left's own chip
    # carry the cue; the Today deadline is bold -- a font-only change, so
    # it can never mask a fill.
    ws.conditional_formatting.add(
        f"{deadline_col}{FIRST_DATA_ROW}:{deadline_col}{LAST_DATA_ROW}",
        FormulaRule(formula=[f"AND({has_deadline},INT({k})=TODAY())"], font=Font(bold=True)),
    )

    deadline_range = f"{deadline_col}{FIRST_DATA_ROW}:{deadline_col}{LAST_DATA_ROW}"
    for fmt_id, offset, word in ((210, 0, "Today"), (211, 1, "Tomorrow"), (212, -1, "Yesterday")):
        rule = Rule(type="expression", formula=[f"AND({has_deadline},INT({k})=TODAY(){offset:+d})"])
        rule.dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=fmt_id, formatCode=f'"{word} "hh:mm'))
        ws.conditional_formatting.add(deadline_range, rule)

    lc = f"${linked_campaign_col}{FIRST_DATA_ROW}"
    ws.conditional_formatting.add(
        f"{linked_campaign_col}{FIRST_DATA_ROW}:{linked_campaign_col}{LAST_DATA_ROW}",
        FormulaRule(formula=[f'{lc}="None"'], font=Font(color=ds.GRIDLINE_DARK)),
    )


def progress_formula(key_col, open_col, done_col, row):
    """ "2/5  ▓▓░░░" -- resolved (completed or skipped) subtasks out of all
    subtasks, at every depth: Total Subtasks (open descendants, a formula
    returning "" for none) + DoneSubtasks (resolved descendants, kept by
    modTasks' archiving cascade). Blank for a row with no subtasks ever.
    Mirrored in modTasks.ProgressFormula for rows the row extender adds.
    """
    done = f"N(${done_col}{row})"
    open_ = f'IF(${open_col}{row}="",0,${open_col}{row})'
    total = f"({done}+{open_})"
    filled = f"ROUND({done}/{total}*5,0)"
    return (
        f'=IF(OR(${key_col}{row}="",{total}=0),"",'
        f'{done}&"/"&{total}&"  "&REPT("\u2593",{filled})&REPT("\u2591",5-{filled}))'
    )


def _apply_progress_complete(ws, progress_col, open_col, done_col):
    """Progress turns green once every subtask is resolved -- the cue that
    the parent itself can be marked Completed."""
    from openpyxl.formatting.rule import FormulaRule

    open_ = f'IF(${open_col}{FIRST_DATA_ROW}="",0,${open_col}{FIRST_DATA_ROW})'
    ws.conditional_formatting.add(
        f"{progress_col}{FIRST_DATA_ROW}:{progress_col}{LAST_DATA_ROW}",
        FormulaRule(
            formula=[f"AND(N(${done_col}{FIRST_DATA_ROW})>0,{open_}=0)"],
            font=Font(color=ds.STATUS_GOOD, bold=True),
        ),
    )

