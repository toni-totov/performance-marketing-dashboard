"""Campaigns -- the user's own personal campaign lifecycle tracker
(PLAN.md §10). This module builds the *structure* only (Phase 2): the
20-slot reserved-block layout, columns, the linked-tasks summary lookup
formulas, the 15-week Gantt, and conditional formatting. Interactive
behaviors (collapse/expand, Status -> Completed archiving) are VBA,
added in Phase 3.

Layout: 20 pre-allocated campaign "slots," each a 10-row block (1 visible
main row + 1 hidden summary-block header row + 8 hidden linked-task
summary rows) -- see §10.2 point 4 for why this can't be a simple
auto-expanding Table.

  Row 1 (hidden): A1 = Gantt rolling-window anchor (live TODAY()-based
    formula, no VBA needed -- see this sheet's design discussion), J1:X1
    = the 15 weeks' raw start dates.
  Row 2 (visible header): A2:I2 = column labels, J2:X2 = week-range
    labels ("06-12 Jul") referencing row 1.
  Rows 3-202: the 20 slots.

Columns: Up (A) / Down (B) -- reorder buttons, added §19 Phase 5, user-
requested, mirroring Tasks' own Up/Down convention (hierarchy.py's
write_id_reorder_columns) -- then Campaign Name, Number of Tasks, Start
Date, End Date, Status, Progress, Notes (C-I), then the 15 Gantt columns
(J-X), then hidden CampaignKey (Y) and CampaignNameList (Z). Unlike
Tasks, there's no hierarchy/SiblingOrder engine here (Campaigns is a
fixed 20-slot layout, not an auto-expanding Table, §10.2 point 4) --
reordering means swapping a campaign's identity fields (Name, Start/End
Date, Status, Notes, CampaignKey) between two main rows, done by VBA
(modCampaigns.bas's HandleCampaignReorder), never moving actual rows.
Skips over empty slots to find the nearest real campaign to swap with
(confirmed with the user) -- a literal adjacent-slot swap would often be
a no-op given campaigns aren't necessarily packed into consecutive slots.

CampaignNameList (Z3:Z22, a small fixed 20-row block right after the
header, one slot per possible campaign) is a plain VALUE-only compacted
list, entirely VBA-maintained (modCampaigns.bas's
RefreshCampaignNameList, called on any Campaign Name edit and on
Workbook_Open) -- NOT a formula mirror, on purpose. **Design correction,
found live (PLAN.md §19 Phase 5):** the original approach was a per-row
formula (`=IF(A{row}="","",A{row})`) mirroring each of the 20 main rows'
Campaign Name into a contiguous range, so Tasks' Linked Campaign dropdown
(a classic List Data Validation) would have a single-range source instead
of the 20 non-contiguous Campaign Name cells directly. That fixed the
cell's own DISPLAY (blank instead of "0"/stray empty string) but not the
dropdown itself: classic Excel List validation only skips a genuinely
untouched blank cell, never a cell holding a FORMULA that merely
evaluates to "" -- so every one of the ~19 typically-unused slots still
showed as a blank, selectable dropdown row. Fixed with a VBA-maintained
plain-value list: every currently-named campaign written contiguously
from the top as literal text.

**Second correction, found live (user still saw one trailing blank,
hoverable/selectable dropdown row even after the above fix):** a FIXED-
size source range (e.g. Z3:Z22, with unused rows past the real count
explicitly `ClearContents`'d) still isn't enough -- Excel's classic List
validation, sourced from a genuinely blank cell that's *within* the
named range but past the real data, still renders one selectable blank
row for it, confirmed live. The real fix is a range that's *exactly* as
tall as the actual entry count, achieved via the classic dynamic-named-
range trick: `CampaignNameList` is defined as
`=OFFSET(Campaigns!$Z$3,0,0,COUNTA(Campaigns!$Z$3:$Z$23),1)` rather than
a fixed A1-style range, so the reference itself grows/shrinks with
COUNTA -- zero rows past the real data are ever included, not just
zeroed-out.

Z3 permanently holds the literal text "None" (written once at build
time, never touched by VBA), reserved as an explicit "not linked to any
campaign" choice -- both because it doubles as the guaranteed-non-blank
floor that keeps COUNTA (and therefore OFFSET's height) always >= 1
(avoiding a zero-height OFFSET, which Excel's List validation shows as
an outright invalid/errored dropdown rather than an empty one), and
because a new task now pre-fills its own Linked Campaign with "None"
(modTasks.bas's HandleTaskNameChange) instead of sitting blank -- no
real campaign is ever literally named "None", so
HandleLinkedCampaignChange's normal "no match found" path naturally
leaves LinkedCampaignKey blank for it, identical to true "unlinked."
Real campaign names are VBA-maintained (RefreshCampaignNameList) starting
at Z4, compacted with no gaps, room for up to NUM_SLOTS (20) of them
(Z4:Z23).

Two more hidden columns, added when Dashboard was built (§4.E, Phase 2J):
ReminderDaysLeft (AA) -- each main row's "days to next boundary" (not yet
started: days until Start Date; already started: days until End Date) --
and ReminderSortKey (AB) -- blank unless the campaign is Planned/Active AND
within the 7-day reminder window (ABS(ReminderDaysLeft)<=7), else
ReminderDaysLeft + CampaignKey/1e6 (unique, sortable). Same SortKey
pattern already used by Tasks (§6, feeding Dashboard §4.B) -- SMALL()
across this sparse column works fine even though the 20 main rows aren't
contiguous, since SMALL() (like MIN()) silently skips blank cells; no new
contiguous mirror needed the way CampaignNameList was for Data Validation.

Each slot's first hidden row is a mini header ("Task Name" / "Subtasks" /
"Deadline" / "Status" / "Notes"), added so the nested list orients itself
without depending on the main row's differently-labeled columns still
being visible/remembered once scrolled. The 8 hidden summary rows below
it (rank N, 1-8) reuse the 5 columns starting at Campaign Name (C-G) for
Task Name (indented) / Number of Subtasks / Deadline / Status text /
Notes icon -- those columns' main-row meanings (Campaign Name/Number of
Tasks/Start Date/End Date/Status) don't apply to a summary row, so
reusing them keeps everything in the same visual columns for a clean
nested-list read when a block is expanded. The Up/Down reorder columns
(A/B) are main-row-only and stay blank on every summary row. Deadline
(not the earlier Days Left) carries the urgency color-coding directly,
computed against TODAY() rather than via a separate stored days-
remaining number.

**Design correction found during Phase 2 implementation (live-tested):**
conditional formatting was originally planned as one rule per Gantt
week-column, using a multi-area (non-contiguous) sqref spanning all 20
campaign rows. Live testing proved Excel does NOT retarget relative
references per sub-area in a multi-area CF range -- every cell in the
union evaluates against the FIRST area's anchor row only. The fix: one CF
rule PER CAMPAIGN SLOT instead (a contiguous row range within that single
slot), looped 20 times. Same fix applies to Status coloring and the
summary block's Days Left urgency coloring -- anywhere a rule would
otherwise need to span multiple non-contiguous rows.
"""

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import conditional_formatting as cf
from . import design_system as ds
from .week_grid import current_week_monday_formula, month_grid_anchor_formula  # noqa: F401 (anchor kept for reference)
from .week_grid import week_label_formula, week_number_formula, week_start_formula

SHEET_NAME = "Campaigns"

HIDDEN_DATE_ROW = 1
HEADER_ROW = 2
FIRST_DATA_ROW = 3
NUM_SLOTS = 20
SUMMARY_ROWS_PER_SLOT = 8
SLOT_SIZE = 2 + SUMMARY_ROWS_PER_SLOT  # 10: main row + header row + 8 data rows
NUM_GANTT_WEEKS = 15

_DATE_FORMAT = "DD.MM.YYYY"
_PERCENT_FORMAT = "0%"

_SUMMARY_HEADER_LABELS = ["Task Name", "Subtasks", "Deadline", "Status", "Notes"]

# Reorder buttons (§19 Phase 5, user-requested) -- always columns A/B,
# ahead of everything else.
_UP_COL = 1
_DOWN_COL = 2
_FIRST_VISIBLE_COL = 3  # Campaign Name starts here

_VISIBLE_HEADERS = ["Campaign Name", "Number of Tasks", "Start Date", "End Date", "Status", "Progress", "Notes"]
# Column indices (1-based) for the 7 visible columns, by name
_COL = {name: i + _FIRST_VISIBLE_COL for i, name in enumerate(_VISIBLE_HEADERS)}
_GANTT_FIRST_COL = _FIRST_VISIBLE_COL + len(_VISIBLE_HEADERS)  # J = 10
_GANTT_LAST_COL = _GANTT_FIRST_COL + NUM_GANTT_WEEKS - 1  # X = 24
_CAMPAIGN_KEY_COL = _GANTT_LAST_COL + 1  # Y = 25
_CAMPAIGN_NAME_LIST_COL = _CAMPAIGN_KEY_COL + 1  # Z = 26
_REMINDER_DAYS_LEFT_COL = _CAMPAIGN_NAME_LIST_COL + 1  # AA = 27
_REMINDER_SORT_KEY_COL = _REMINDER_DAYS_LEFT_COL + 1  # AB = 28

# Tasks' own column letters this sheet reads from (see sheets/tasks.py)
_TASKS_TASK_COL = "E"
_TASKS_NUM_SUBTASKS_COL = "F"
_TASKS_DEADLINE_COL = "K"
_TASKS_STATUS_COL = "L"
_TASKS_NOTES_COL = "P"
_TASKS_CAMPAIGN_RANK_KEY_COL = "Z"
_TASKS_LINKED_CAMPAIGN_KEY_COL = "V"
_TASKS_LAST_ROW = 501
_TASKS_FIRST_ROW = 2


def _slot_main_row(slot_index_0based):
    return FIRST_DATA_ROW + slot_index_0based * SLOT_SIZE


def build(wb):
    """Create the Campaigns sheet in `wb` (an openpyxl Workbook). Returns
    the created worksheet. Requires Tasks to already exist in `wb`.
    """
    ws = wb.create_sheet(SHEET_NAME)
    ds.apply_modern_chrome(ws)
    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=1)

    # --- Row 1 (hidden): Gantt anchor + raw week-start dates ---
    ws.cell(row=HIDDEN_DATE_ROW, column=1, value=current_week_monday_formula())
    anchor_ref = f"A{HIDDEN_DATE_ROW}"
    for w in range(NUM_GANTT_WEEKS):
        col = _GANTT_FIRST_COL + w
        ws.cell(row=HIDDEN_DATE_ROW, column=col, value=week_start_formula(anchor_ref, w))
    ws.row_dimensions[HIDDEN_DATE_ROW].hidden = True

    # --- Row 2 (visible header) ---
    for col, label in ((_UP_COL, "↑"), (_DOWN_COL, "↓")):
        cell = ws.cell(row=HEADER_ROW, column=col, value=label)
        ds.style_header_cell(cell)

    for i, label in enumerate(_VISIBLE_HEADERS, start=_FIRST_VISIBLE_COL):
        cell = ws.cell(row=HEADER_ROW, column=i, value=label)
        ds.style_header_cell(cell)

    # Gantt week headers: "Week NN" + the date-range label, combined into
    # one wrapped, two-line cell (user-requested, PLAN.md §19 Phase 5) --
    # a real second header row would mean shifting every row on the sheet
    # down by one, touching row-position math in campaigns.py,
    # modCampaigns.bas, modTasks.bas/modRecurringTasks.bas, and
    # dashboard.py for a purely cosmetic gain; confirmed with the user
    # that a taller single wrapped cell achieves the same readability
    # with none of that risk.
    ws.row_dimensions[HEADER_ROW].height = ds.HEADER_ROW_HEIGHT
    gap_border = Border(left=Side(style="medium", color=ds.INK_MUTED))
    for w in range(NUM_GANTT_WEEKS):
        col = _GANTT_FIRST_COL + w
        week_start_ref = f"{get_column_letter(col)}{HIDDEN_DATE_ROW}"
        week_end_ref = f"{week_start_ref}+6"
        week_number_expr = week_number_formula(week_start_ref)
        range_label_expr = week_label_formula(week_start_ref, week_end_ref)[1:]  # strip leading "="
        cell = ws.cell(row=HEADER_ROW, column=col, value=f'="Week "&{week_number_expr}&CHAR(10)&{range_label_expr}')
        ds.style_header_cell(cell)
        if w == 0:
            # Visual gap between the campaign table and the Gantt (user-
            # requested) -- a medium left border on the boundary column,
            # rather than inserting an actual spacer column (another
            # cross-file column shift for a purely visual gain).
            cell.border = gap_border

    ws.cell(row=HEADER_ROW, column=_CAMPAIGN_KEY_COL, value="CampaignKey")
    ws.cell(row=HEADER_ROW, column=_CAMPAIGN_NAME_LIST_COL, value="CampaignNameList")
    ws.cell(row=HEADER_ROW, column=_REMINDER_DAYS_LEFT_COL, value="ReminderDaysLeft")
    ws.cell(row=HEADER_ROW, column=_REMINDER_SORT_KEY_COL, value="ReminderSortKey")

    # --- Column widths ---
    widths = {
        get_column_letter(_UP_COL): 4,
        get_column_letter(_DOWN_COL): 4,
        get_column_letter(_COL["Campaign Name"]): 22,
        get_column_letter(_COL["Number of Tasks"]): 12,
        get_column_letter(_COL["Start Date"]): 13,
        get_column_letter(_COL["End Date"]): 13,
        get_column_letter(_COL["Status"]): 11,
        get_column_letter(_COL["Progress"]): 10,
        get_column_letter(_COL["Notes"]): 8,
    }
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width
    for w in range(NUM_GANTT_WEEKS):
        ws.column_dimensions[get_column_letter(_GANTT_FIRST_COL + w)].width = 9
    ws.column_dimensions[get_column_letter(_CAMPAIGN_KEY_COL)].hidden = True
    ws.column_dimensions[get_column_letter(_CAMPAIGN_NAME_LIST_COL)].hidden = True
    ws.column_dimensions[get_column_letter(_REMINDER_DAYS_LEFT_COL)].hidden = True
    ws.column_dimensions[get_column_letter(_REMINDER_SORT_KEY_COL)].hidden = True

    # --- CampaignNameList's permanent "None" floor (see module docstring)
    # -- written once here, never touched by RefreshCampaignNameList,
    # which only manages X(FIRST_DATA_ROW+1) onward. ---
    name_list_letter = get_column_letter(_CAMPAIGN_NAME_LIST_COL)
    ws.cell(row=FIRST_DATA_ROW, column=_CAMPAIGN_NAME_LIST_COL, value="None")

    # --- Per-slot content ---
    for slot in range(NUM_SLOTS):
        main_row = _slot_main_row(slot)
        _build_main_row(ws, main_row)
        _build_summary_rows(ws, main_row)
        # Row structure (§17.12 correction, user-reported) -- table
        # columns only (1.._GANTT_FIRST_COL-1), not the Gantt itself:
        # _build_main_row() above already set a left-edge gap border on
        # this row's _GANTT_FIRST_COL cell (the table/Gantt visual
        # boundary), and this call would silently overwrite it if it
        # included that column too. Applied BEFORE the button styling
        # below so each button cell's own border wins on top of this
        # row's bottom-only one.
        ds.apply_table_hairline_borders(ws, 1, _GANTT_FIRST_COL - 1, main_row, main_row)
        # Button-shaped up/down cells (§17.11) are deliberately NOT
        # applied here at build time -- §17.13's correction, same
        # reasoning as Tasks. Applied dynamically instead, only to
        # occupied slots, by modCampaigns.bas's CompactCampaignSlots.

    # --- Named range for the Linked Campaign dropdown source (Tasks
    # references this by name -- see tasks.py). A dynamic OFFSET/COUNTA
    # reference, not a fixed range -- see module docstring for why a
    # fixed range (even with unused rows explicitly cleared to real
    # blanks) still left one trailing selectable blank row in the actual
    # dropdown UI. ---
    from openpyxl.workbook.defined_name import DefinedName

    name_list_last_row = FIRST_DATA_ROW + NUM_SLOTS
    name_list_range = f"'{SHEET_NAME}'!${name_list_letter}${FIRST_DATA_ROW}:${name_list_letter}${name_list_last_row}"
    wb.defined_names["CampaignNameList"] = DefinedName(
        "CampaignNameList",
        attr_text=(
            f"OFFSET('{SHEET_NAME}'!${name_list_letter}${FIRST_DATA_ROW},0,0,"
            f"COUNTA({name_list_range}),1)"
        ),
    )

    ws.sheet_properties.tabColor = ds.TAB_CORE

    return ws


def _build_main_row(ws, row):
    key_letter = get_column_letter(_CAMPAIGN_KEY_COL)
    name_list_letter = get_column_letter(_CAMPAIGN_NAME_LIST_COL)

    # Campaign Name: blank, user-typed. Number of Tasks: live count keyed
    # off Tasks!LinkedCampaignKey (§10.2 point 2). Blank instead of a
    # literal "0" when no tasks are linked yet -- same convention as
    # Number of Subtasks (see hierarchy.py's number_of_subtasks_formula).
    count_expr = (
        f"COUNTIFS(Tasks!${_TASKS_LINKED_CAMPAIGN_KEY_COL}${_TASKS_FIRST_ROW}:"
        f"${_TASKS_LINKED_CAMPAIGN_KEY_COL}${_TASKS_LAST_ROW},{key_letter}{row})"
    )
    center = Alignment(horizontal="center", vertical="center")
    num_tasks_cell = ws.cell(row=row, column=_COL["Number of Tasks"])
    num_tasks_cell.value = f'=IF({count_expr}=0,"",{count_expr})'
    num_tasks_cell.alignment = center  # §17.14 correction (user-reported)

    start_col_letter = get_column_letter(_COL["Start Date"])
    end_col_letter = get_column_letter(_COL["End Date"])
    ws.cell(row=row, column=_COL["Start Date"]).number_format = _DATE_FORMAT
    ws.cell(row=row, column=_COL["End Date"]).number_format = _DATE_FORMAT

    # Progress (§10.1): 0-100% elapsed time between Start/End vs. today;
    # single-day campaigns read 0% before that day, 100% on/after.
    s, e = f"{start_col_letter}{row}", f"{end_col_letter}{row}"
    progress_cell = ws.cell(row=row, column=_COL["Progress"])
    progress_cell.value = (
        f'=IF(OR({s}="",{e}=""),"",IF({s}={e},IF(TODAY()>={s},1,0),'
        f"MAX(0,MIN(1,(TODAY()-{s})/({e}-{s})))))"
    )
    progress_cell.number_format = _PERCENT_FORMAT
    progress_cell.alignment = center  # §17.14 correction (user-reported)

    # Sequential blue ramp per §17.5 (PLAN.md §19 Phase 6 -- confirmed
    # with the user this was missing entirely, added now). Applied
    # per-slot, same as the Status rule above -- a combined multi-area
    # rule across all 20 slots doesn't retarget correctly (see this
    # function's earlier note). Same light 100-300 range as Tasks' depth
    # shading, not the ramp's darker steps, so INK_PRIMARY text stays
    # readable on top.
    progress_letter = get_column_letter(_COL["Progress"])
    cf.apply_rules(
        ws,
        cf.sequential_rules(
            progress_letter, row, row, buckets=[(0, 100), (0.2, 150), (0.4, 200), (0.6, 250), (0.8, 300)]
        ),
    )

    notes_cell = ws.cell(row=row, column=_COL["Notes"])
    notes_cell.number_format = ds.notes_icon_format()
    _block_direct_entry(ws, get_column_letter(_COL["Notes"]), row, "Double-click this cell to open the Notes popup instead.")

    status_dv = DataValidation(type="list", formula1='"Planned,Active,Completed"', allow_blank=True)
    status_col_letter = get_column_letter(_COL["Status"])
    status_dv.add(f"{status_col_letter}{row}")
    ws.add_data_validation(status_dv)

    # CampaignNameList (X3:X22) is intentionally NOT written here -- see
    # module docstring: it's a VBA-maintained plain-value compacted list,
    # not a per-row formula mirror, rebuilt by
    # modCampaigns.RefreshCampaignNameList whenever any Campaign Name
    # changes (and once on Workbook_Open, to backfill pre-existing data).

    # --- ReminderDaysLeft / ReminderSortKey (§4.E, Dashboard) -- see
    # module docstring. Only written on main rows, same as
    # CampaignNameList; summary rows' cells stay genuinely blank. ---
    key_cell = f"{key_letter}{row}"
    days_left_cell = f"{get_column_letter(_REMINDER_DAYS_LEFT_COL)}{row}"
    ws.cell(row=row, column=_REMINDER_DAYS_LEFT_COL).value = f"=IF(TODAY()<{s},{s}-TODAY(),{e}-TODAY())"
    ws.cell(row=row, column=_REMINDER_SORT_KEY_COL).value = (
        f'=IF(AND(OR({status_col_letter}{row}="Planned",{status_col_letter}{row}="Active"),'
        f"ABS({days_left_cell})<=7),{days_left_cell}+{key_cell}/1000000,\"\")"
    )

    # --- Conditional formatting, per-slot (see module docstring: a
    # combined multi-area rule across all 20 slots would NOT retarget
    # correctly, confirmed via live testing) ---
    cf.apply_rules(
        ws,
        cf.workflow_status_rules(
            status_col_letter, row, row, not_started_label="Planned", in_motion_label="Active"
        ),
    )

    gantt_range = f"{get_column_letter(_GANTT_FIRST_COL)}{row}:{get_column_letter(_GANTT_LAST_COL)}{row}"
    from openpyxl.formatting.rule import FormulaRule

    # Visual gap between the campaign table and the Gantt (user-
    # requested) -- same left border as the header row's first Gantt
    # cell, applied here too so the boundary reads consistently down the
    # whole column, not just at the top.
    ws.cell(row=row, column=_GANTT_FIRST_COL).border = Border(left=Side(style="medium", color=ds.INK_MUTED))

    # **Bug found live (PLAN.md §19 Phase 5, user-reported):** Start/End
    # Date ($s$/$e$ below) were NOT column-locked, but this rule's range
    # spans 15 COLUMNS in a single row -- Excel shifts every unlocked
    # reference sideways as it evaluates each column of the range, so
    # Start/End Date's reference literally drifted into whatever cell
    # sat N columns to the right of it, for the Nth Gantt column. That's
    # also why cells could show colored before any date was even entered:
    # the drifted reference was landing on some unrelated cell that
    # happened to satisfy the condition. Confirmed by comparing against
    # Monthly Media Plan's own (correct) Gantt CF, which locks the
    # opposite axis because its range spans rows within one column
    # instead. Fixed with $-locked columns on Start/End Date here; the
    # week-anchor reference is correctly left column-UNlocked (only the
    # row is locked), since that's what makes it retarget to each
    # column's own week as the rule evaluates across the row.
    s_locked, e_locked = f"${start_col_letter}{row}", f"${end_col_letter}{row}"
    status_locked = f"${status_col_letter}{row}"
    week1_start_ref = f"{get_column_letter(_GANTT_FIRST_COL)}${HIDDEN_DATE_ROW}"
    overlap_condition = f"AND({s_locked}<={week1_start_ref}+6,{e_locked}>={week1_start_ref})"

    # User-requested addition: color the bar by the campaign's own
    # Status (same 3-state workflow palette used everywhere else, §17.4)
    # instead of one flat brand-accent color -- shows at a glance whether
    # a bar is still Planned vs. already Active.
    # bgColor, not fgColor -- see conditional_formatting.py's _fill() for
    # why (a differential/CF PatternFill reads its displayed color from
    # bgColor, confirmed via live COM testing during the Monthly Media
    # Plan phase; this hand-rolled rule bypasses that shared helper so it
    # needed the same fix applied directly).
    for status_label, color in (
        ("Planned", ds.WORKFLOW_NOT_STARTED),
        ("Active", ds.WORKFLOW_IN_MOTION),
        ("Completed", ds.WORKFLOW_DONE),
    ):
        condition = f'AND({overlap_condition},{status_locked}="{status_label}")'
        ws.conditional_formatting.add(gantt_range, FormulaRule(formula=[condition], fill=PatternFill("solid", bgColor=color)))


def _build_summary_rows(ws, main_row):
    """Writes this slot's mini header row (main_row+1) and its 8 hidden
    linked-task data rows (main_row+2 .. main_row+9). See module
    docstring for the column layout (Task Name/Subtasks/Deadline/Status/
    Notes, reusing the 5 columns starting at Campaign Name) and why
    Deadline (not a separate Days Left number) carries the urgency
    coloring.
    """
    header_row = main_row + 1

    # Reused main-row columns, in summary-row order (Task Name/Subtasks/
    # Deadline/Status/Notes) -- see module docstring.
    task_name_col = _COL["Campaign Name"]
    num_subtasks_col = _COL["Number of Tasks"]
    deadline_col = _COL["Start Date"]
    status_col = _COL["End Date"]
    notes_col = _COL["Status"]

    header_font = Font(name=ds.FONT_NAME, size=ds.FONT_SIZE_CAPTION, bold=True, color=ds.INK_MUTED)
    header_border = Border(bottom=Side(style="thin", color=ds.GRIDLINE))
    for idx, label in enumerate(_SUMMARY_HEADER_LABELS):
        cell = ws.cell(row=header_row, column=task_name_col + idx, value=label)
        cell.font = header_font
        cell.border = header_border
        cell.alignment = Alignment(horizontal="left" if idx == 0 else "center")
    ws.row_dimensions[header_row].hidden = True

    key_letter = get_column_letter(_CAMPAIGN_KEY_COL)
    key_cell = f"${key_letter}{main_row}"
    rank_key_range = (
        f"Tasks!${_TASKS_CAMPAIGN_RANK_KEY_COL}${_TASKS_FIRST_ROW}:"
        f"${_TASKS_CAMPAIGN_RANK_KEY_COL}${_TASKS_LAST_ROW}"
    )
    task_name_range = f"Tasks!${_TASKS_TASK_COL}${_TASKS_FIRST_ROW}:${_TASKS_TASK_COL}${_TASKS_LAST_ROW}"
    num_subtasks_range = (
        f"Tasks!${_TASKS_NUM_SUBTASKS_COL}${_TASKS_FIRST_ROW}:${_TASKS_NUM_SUBTASKS_COL}${_TASKS_LAST_ROW}"
    )
    deadline_range = f"Tasks!${_TASKS_DEADLINE_COL}${_TASKS_FIRST_ROW}:${_TASKS_DEADLINE_COL}${_TASKS_LAST_ROW}"
    status_range = f"Tasks!${_TASKS_STATUS_COL}${_TASKS_FIRST_ROW}:${_TASKS_STATUS_COL}${_TASKS_LAST_ROW}"
    notes_range = f"Tasks!${_TASKS_NOTES_COL}${_TASKS_FIRST_ROW}:${_TASKS_NOTES_COL}${_TASKS_LAST_ROW}"
    num_tasks_cell = f"${get_column_letter(_COL['Number of Tasks'])}{main_row}"
    # N(...) coerces Number of Tasks' now-possibly-blank ("" when zero,
    # see _build_main_row) value to 0 -- without it, Excel's text->number
    # comparison rules make "">8 evaluate TRUE (text always sorts above
    # numbers), wrongly triggering overflow logic for a zero-task
    # campaign.
    num_tasks_n = f"N({num_tasks_cell})"

    for n in range(1, SUMMARY_ROWS_PER_SLOT + 1):
        row = header_row + n
        target_key = f'{key_cell}&"|"&{n}'
        match_expr = f"MATCH({target_key},{rank_key_range},0)"
        is_overflow_row = n == SUMMARY_ROWS_PER_SLOT

        task_name_formula = f'=IFERROR(INDEX({task_name_range},{match_expr}),"")'
        if is_overflow_row:
            # Overflow row: show "+N more" instead of the 8th task if
            # there are more than 8 linked tasks.
            task_name_formula = (
                f'=IF({num_tasks_n}>{SUMMARY_ROWS_PER_SLOT},'
                f'"+ "&({num_tasks_n}-{SUMMARY_ROWS_PER_SLOT - 1})&" more — see Tasks",'
                f'{task_name_formula[1:]})'
            )
        ws.cell(row=row, column=task_name_col).value = task_name_formula
        ws.cell(row=row, column=task_name_col).alignment = Alignment(indent=1)
        ws.cell(row=row, column=task_name_col).font = Font(name=ds.FONT_NAME, size=ds.FONT_SIZE_BODY, italic=True, color=ds.INK_SECONDARY)

        # Number of Subtasks -- blank on the overflow row (same guard as
        # the Task Name cell above) since that row shows "+N more" text
        # instead of a single real task.
        num_subtasks_formula = f'=IFERROR(INDEX({num_subtasks_range},{match_expr}),"")'
        if is_overflow_row:
            num_subtasks_formula = f'=IF({num_tasks_n}>{SUMMARY_ROWS_PER_SLOT},"",{num_subtasks_formula[1:]})'
        ws.cell(row=row, column=num_subtasks_col).value = num_subtasks_formula
        ws.cell(row=row, column=num_subtasks_col).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=num_subtasks_col).font = Font(name=ds.FONT_NAME, size=ds.FONT_SIZE_BODY, color=ds.INK_SECONDARY)

        # Deadline -- carries urgency coloring directly (see
        # cf.deadline_urgency_rules below), replacing the earlier
        # separate Days Left number. INDEX() on a genuinely blank source
        # cell (a matched task with no Deadline set yet) returns numeric
        # 0, not blank -- under a date format that displays as "00.01.1900"
        # (Excel's day-0 epoch), not empty. The explicit "=0" check catches
        # that (a real deadline is never legitimately day 0).
        deadline_idx = f"INDEX({deadline_range},{match_expr})"
        deadline_formula = f'=IFERROR(IF({deadline_idx}=0,"",{deadline_idx}),"")'
        if is_overflow_row:
            deadline_formula = f'=IF({num_tasks_n}>{SUMMARY_ROWS_PER_SLOT},"",{deadline_formula[1:]})'
        ws.cell(row=row, column=deadline_col).value = deadline_formula
        ws.cell(row=row, column=deadline_col).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=deadline_col).number_format = _DATE_FORMAT

        # Status: plain text ("Not Started"/"In Progress" -- only two
        # possible states for an open Task, §10.2), not an icon (user
        # feedback, PLAN.md §19 Phase 5: the "Not Started" icon, a hollow
        # white square, read as an empty checkbox rather than a status
        # indicator, especially now that every new task defaults to
        # "Not Started"). Same INDEX-on-blank-returns-0 guard as
        # Deadline/Notes above.
        status_idx = f"INDEX({status_range},{match_expr})"
        status_formula = f'=IFERROR(IF({status_idx}=0,"",{status_idx}),"")'
        if is_overflow_row:
            status_formula = f'=IF({num_tasks_n}>{SUMMARY_ROWS_PER_SLOT},"",{status_formula[1:]})'
        ws.cell(row=row, column=status_col).value = status_formula
        ws.cell(row=row, column=status_col).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=status_col).font = Font(name=ds.FONT_NAME, size=ds.FONT_SIZE_BODY, color=ds.INK_SECONDARY)

        # Notes: masked icon format, same convention as every other
        # Notes column in the workbook (ds.notes_icon_format()) -- the
        # cell's real stored value is the note text itself, only its
        # DISPLAY is replaced by an icon when non-blank. Same INDEX-on-
        # blank-returns-0 issue as Deadline above, but the "zero" section
        # of notes_icon_format() is "General" (matching a normal Notes
        # column, where a stray 0 never occurs) -- so a task with no note
        # showed a literal "0" instead of nothing. Same "=0" guard fixes
        # it.
        notes_idx = f"INDEX({notes_range},{match_expr})"
        notes_formula = f'=IFERROR(IF({notes_idx}=0,"",{notes_idx}),"")'
        if is_overflow_row:
            notes_formula = f'=IF({num_tasks_n}>{SUMMARY_ROWS_PER_SLOT},"",{notes_formula[1:]})'
        ws.cell(row=row, column=notes_col).value = notes_formula
        ws.cell(row=row, column=notes_col).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=notes_col).number_format = ds.notes_icon_format()

        ws.row_dimensions[row].hidden = True

    # Deadline urgency coloring + Status workflow coloring + subtask-
    # prefix marker across this slot's 8 (contiguous) data rows --
    # per-slot, same reasoning as the Gantt/Status CF above. Excludes the
    # header row (header_row itself), only header_row+1 .. header_row+8.
    deadline_letter = get_column_letter(deadline_col)
    cf.apply_rules(
        ws, cf.deadline_urgency_rules(deadline_letter, header_row + 1, header_row + SUMMARY_ROWS_PER_SLOT)
    )

    # User-requested addition (PLAN.md §19 Phase 5): color-code Status the
    # same way Tasks' own Status column and Campaigns' main-row Status
    # already are (§17.4's 3-state scheme) -- default labels match
    # exactly, since this column pulls Tasks' own "Not Started"/
    # "In Progress" text verbatim. "Completed" never actually occurs here
    # (only open tasks are linked, §10.2), so that 3rd rule is inert but
    # harmless.
    status_letter = get_column_letter(status_col)
    cf.apply_rules(
        ws, cf.workflow_status_rules(status_letter, header_row + 1, header_row + SUMMARY_ROWS_PER_SLOT)
    )

    _apply_subtask_prefix(ws, header_row)


def _apply_subtask_prefix(ws, header_row):
    """User-requested addition (PLAN.md §19 Phase 5): same "⤷ " marker
    used on Tasks/Recurring Tasks (see tasks.py's _apply_subtask_prefix),
    reused here so a campaign's linked tasks visually read as a nested
    list once the block is expanded. Unlike Tasks/Recurring Tasks there is
    no hierarchy Level column to key off of here -- every data row IS a
    subtask row by construction, so the condition is simply "this row has
    any displayed text" (covers both a real linked-task name and the
    "+N more" overflow row), which also means a genuinely empty/unused
    slot (IFERROR(...,"") fallback) correctly stays unmarked. Deliberately
    excludes the mini header row directly above these data rows -- its
    "Task Name" label would otherwise also match the same non-blank
    condition and get a nonsensical "⤷ Task Name" prefix.
    """
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.styles.numbers import NumberFormat

    name_col_letter = get_column_letter(_COL["Campaign Name"])
    task_range = f"{name_col_letter}{header_row + 1}:{name_col_letter}{header_row + SUMMARY_ROWS_PER_SLOT}"
    cell = f"${name_col_letter}{header_row + 1}"

    rule = Rule(type="expression", formula=[f'{cell}<>""'])
    # Distinct numFmtId from Tasks' (201) and Recurring Tasks' (202) --
    # custom formats are workbook-scoped, so IDs must not collide.
    rule.dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=203, formatCode='"⤷ "@'))
    ws.conditional_formatting.add(task_range, rule)


def _block_direct_entry(ws, column_letter, row, message):
    dv = DataValidation(
        type="custom", formula1="FALSE", showErrorMessage=True, errorTitle="Use the popup instead", error=message
    )
    dv.add(f"{column_letter}{row}")
    ws.add_data_validation(dv)
