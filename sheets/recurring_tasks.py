"""Recurring Tasks -- templates for Daily/Weekly/Monthly tasks that, when
Active, automatically generate real rows into Tasks on their schedule
(PLAN.md §12). This module builds the *structure* only (Phase 2): columns,
formatting, conditional formatting, data validation, and the
`WeekdayList`/`DayNumberList` named ranges the Task Start On dropdown
depends on. `RegenerateRecurringTasks()` (the watermark-based generation
macro, §12.1) and the hierarchy engine's runtime behavior are VBA, added
in Phase 3.

Column layout (§12.4), left to right -- Status prepended, otherwise the
same contiguous ID/reorder/Task block Tasks uses (§6.1), followed by
Recurring Tasks' own columns in place of Tasks' Start Date/Deadline/Days
Left/Timer/Time Spent (none of which apply to a template, §12.4):
  Status, ID, up, down, Task, Number of Subtasks, Subtask of,
  Project/Category, Linked Campaign, Priority, Recurrence, Task Start On,
  Deadline Offset (days), Notes
Hidden columns, grouped together after Notes: RecurringKey,
RecurringParentKey, Level, SiblingOrder, LinkedCampaignKey. "Last
Generated" (column T, right after LinkedCampaignKey) is NOT hidden --
**design correction, user-requested (PLAN.md §19 Phase 5):** originally
planned as hidden internal state like the others, but unlike
RecurringKey/ParentKey (pure plumbing), it's genuinely useful
information at a glance -- when a template last actually fired -- so
it's a real visible column instead. Root-level rows only (not
structurally enforceable, just semantic; Phase 3's generation macro
only ever reads/writes it on root rows), so it reads blank on subtask
rows.
Further right, a hidden helper block (not a real sheet column, no
header) holds the `WeekdayList` (7 cells) and `DayNumberList` (31 cells)
values that back Task Start On's dynamic INDIRECT-sourced dropdown
(§12.3) -- same "off-to-the-side hidden state" convention as Calendar's
and Monthly Media Plan's own hidden state columns.

This sheet gets its own independent hierarchy tree (own key/ID sequence,
unrelated to Tasks' own), built the same way (§12.2) -- a plain
pre-formatted range, not a native Excel Table, since VBA's
RenumberHierarchy()-equivalent rewrites rows wholesale in place, same
reasoning as Tasks (see tasks.py's own docstring). Pre-allocated to 100
rows (confirmed with the user) rather than Tasks' 500 -- recurring
templates are naturally a much shorter, curated list than actual tasks.

`LinkedCampaignKey` is declared here but left without a formula, same as
Tasks' own column of the same name -- it's resolved by VBA on change
(Phase 3), not a live formula, consistent with the existing precedent.
"""

from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from . import conditional_formatting as cf
from . import design_system as ds
from .hierarchy import (
    number_of_subtasks_formula,
    write_hidden_hierarchy_columns,
    write_id_reorder_columns,
    write_task_columns,
)

SHEET_NAME = "Recurring Tasks"
HEADER_ROW = 1
FIRST_DATA_ROW = 2
LAST_DATA_ROW = 101  # 100 pre-formatted rows, confirmed with the user

_DATE_FORMAT = "DD.MM.YYYY"
_DAYS_FORMAT = "0"

_WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
_WEEKDAY_LIST_COL = 22  # V, hidden helper block (not a real sheet column)
_DAY_NUMBER_LIST_COL = 23  # W


def _style_header_cell(cell):
    ds.style_header_cell(cell)


def _write_header(ws, col, label):
    cell = ws.cell(row=HEADER_ROW, column=col, value=label)
    _style_header_cell(cell)
    return get_column_letter(col)


def _block_direct_entry(ws, column_letter, message):
    dv = DataValidation(
        type="custom",
        formula1="FALSE",
        showErrorMessage=True,
        errorTitle="Use the popup instead",
        error=message,
    )
    dv.add(f"{column_letter}{FIRST_DATA_ROW}:{column_letter}{LAST_DATA_ROW}")
    ws.add_data_validation(dv)


def build(wb):
    """Create the Recurring Tasks sheet in `wb` (an openpyxl Workbook).
    Requires Campaigns to already exist (for the Linked Campaign
    dropdown's CampaignNameList source). Returns the created worksheet.
    """
    ws = wb.create_sheet(SHEET_NAME)
    ds.apply_modern_chrome(ws)
    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=1)
    ws.row_dimensions[HEADER_ROW].height = ds.HEADER_ROW_HEIGHT

    # --- Status (column A), prepended per §12.4 ---
    status_col = _write_header(ws, 1, "Status")

    # --- ID / up / down (columns B-D) ---
    id_cols = write_id_reorder_columns(ws, start_col=2, header_row=HEADER_ROW)

    # --- Task / Number of Subtasks / Subtask of (columns E-G) ---
    task_cols = write_task_columns(ws, start_col=5, header_row=HEADER_ROW)

    # --- Remaining visible columns (H onward) ---
    project_col = _write_header(ws, 8, "Project / Category")
    linked_campaign_col = _write_header(ws, 9, "Linked Campaign")
    priority_col = _write_header(ws, 10, "Priority")
    recurrence_col = _write_header(ws, 11, "Recurrence")
    task_start_on_col = _write_header(ws, 12, "Task Start On")
    deadline_offset_col = _write_header(ws, 13, "Deadline Offset (days)")
    notes_col = _write_header(ws, 14, "Notes")

    # --- Hidden columns, grouped together starting at column O ---
    hidden_cols = write_hidden_hierarchy_columns(
        ws, start_col=15, key_col_name="RecurringKey", parent_key_col_name="RecurringParentKey", header_row=HEADER_ROW
    )
    linked_campaign_key_col = _write_header(ws, 19, "LinkedCampaignKey")
    ws.column_dimensions[linked_campaign_key_col].hidden = True

    # "Last Generated" (column T) -- user-requested (PLAN.md §19 Phase 5):
    # unlike RecurringKey/ParentKey (pure internal plumbing, stay hidden),
    # this is genuinely useful at a glance -- when a template last fired
    # -- so it's a real visible column, not hidden like the rest of this
    # block. Root-level rows only (semantic, not structurally enforced --
    # see module docstring), so it reads blank on subtask rows.
    last_generated_col = _write_header(ws, 20, "Last Generated")

    # --- Column widths (visible columns only) ---
    widths = {
        status_col: 13,
        id_cols["id"]: 10,
        id_cols["up"]: 4,
        id_cols["down"]: 4,
        task_cols["task"]: 32,
        task_cols["num_subtasks"]: 10,
        task_cols["subtask_of"]: 10,
        project_col: 18,
        linked_campaign_col: 18,
        priority_col: 10,
        recurrence_col: 12,
        task_start_on_col: 14,
        deadline_offset_col: 14,
        notes_col: 8,
        last_generated_col: 14,
    }
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width

    # --- Number formats + static formulas across the data range ---
    for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
        ws.cell(row=row, column=13).number_format = _DAYS_FORMAT  # Deadline Offset
        ws.cell(row=row, column=14).number_format = ds.notes_icon_format()  # Notes
        ws.cell(row=row, column=20).number_format = _DATE_FORMAT  # LastGeneratedThroughDate

        num_sub_cell = ws.cell(row=row, column=6)
        num_sub_cell.value = number_of_subtasks_formula(id_cols["id"], row, FIRST_DATA_ROW, LAST_DATA_ROW)

    # --- Data Validation dropdowns ---
    status_dv = DataValidation(type="list", formula1='"Active,Inactive"', allow_blank=True)
    status_dv.add(f"{status_col}{FIRST_DATA_ROW}:{status_col}{LAST_DATA_ROW}")
    ws.add_data_validation(status_dv)

    priority_dv = DataValidation(type="list", formula1='"Low,Medium,High"', allow_blank=True)
    priority_dv.add(f"{priority_col}{FIRST_DATA_ROW}:{priority_col}{LAST_DATA_ROW}")
    ws.add_data_validation(priority_dv)

    recurrence_dv = DataValidation(type="list", formula1='"Daily,Weekly,Monthly"', allow_blank=True)
    recurrence_dv.add(f"{recurrence_col}{FIRST_DATA_ROW}:{recurrence_col}{LAST_DATA_ROW}")
    ws.add_data_validation(recurrence_dv)

    # Linked Campaign, wired the same way as Tasks' (2E) -- CampaignNameList
    # is a contiguous mirror of Campaigns' 20 (non-contiguous) slot cells.
    linked_campaign_dv = DataValidation(type="list", formula1="=CampaignNameList", allow_blank=True)
    linked_campaign_dv.add(f"{linked_campaign_col}{FIRST_DATA_ROW}:{linked_campaign_col}{LAST_DATA_ROW}")
    ws.add_data_validation(linked_campaign_dv)

    # Task Start On (§12.3): dynamic dropdown via Data Validation formula,
    # no VBA needed to swap options as Recurrence changes. A single
    # contiguous range, so the row-relative Recurrence reference retargets
    # correctly per row (same reasoning already confirmed for CF ranges).
    recurrence_cell = f"${recurrence_col}{FIRST_DATA_ROW}"
    task_start_on_dv = DataValidation(
        type="list",
        formula1=(
            f'=INDIRECT(IF({recurrence_cell}="Weekly","WeekdayList",'
            f'IF({recurrence_cell}="Monthly","DayNumberList","")))'
        ),
        allow_blank=True,
    )
    task_start_on_dv.add(f"{task_start_on_col}{FIRST_DATA_ROW}:{task_start_on_col}{LAST_DATA_ROW}")
    ws.add_data_validation(task_start_on_dv)

    # Notes: "not directly typed" (§12.4, same as Tasks §6.1/§6.3) -- set
    # only via the frmNotes popup.
    _block_direct_entry(ws, notes_col, "Double-click this cell to open the Notes popup instead.")

    # --- WeekdayList / DayNumberList hidden helper block (§12.3) ---
    weekday_letter = get_column_letter(_WEEKDAY_LIST_COL)
    for i, name in enumerate(_WEEKDAY_NAMES, start=1):
        ws.cell(row=i, column=_WEEKDAY_LIST_COL, value=name)
    day_number_letter = get_column_letter(_DAY_NUMBER_LIST_COL)
    for day in range(1, 32):
        ws.cell(row=day, column=_DAY_NUMBER_LIST_COL, value=day)
    ws.column_dimensions[weekday_letter].hidden = True
    ws.column_dimensions[day_number_letter].hidden = True

    wb.defined_names["WeekdayList"] = DefinedName(
        "WeekdayList", attr_text=f"'{SHEET_NAME}'!${weekday_letter}$1:${weekday_letter}${len(_WEEKDAY_NAMES)}"
    )
    wb.defined_names["DayNumberList"] = DefinedName(
        "DayNumberList", attr_text=f"'{SHEET_NAME}'!${day_number_letter}$1:${day_number_letter}$31"
    )

    # --- Conditional formatting ---
    cf.apply_rules(ws, cf.binary_state_rules(status_col, FIRST_DATA_ROW, LAST_DATA_ROW, "Active", "Inactive"))
    cf.apply_rules(ws, cf.priority_rules(priority_col, FIRST_DATA_ROW, LAST_DATA_ROW))
    cf.apply_rules(
        ws,
        cf.categorical_rules(
            recurrence_col,
            FIRST_DATA_ROW,
            LAST_DATA_ROW,
            [("Daily", ds.CATEGORICAL[0]), ("Weekly", ds.CATEGORICAL[1]), ("Monthly", ds.CATEGORICAL[2])],
        ),
    )
    _apply_depth_shading(ws, hidden_cols["level"])
    _apply_subtask_prefix(ws, hidden_cols["level"], task_cols["task"])

    # Row structure (§17.12 correction, user-reported).
    ds.apply_table_hairline_borders(ws, 1, 20, FIRST_DATA_ROW, LAST_DATA_ROW)  # through Last Generated

    # Button-shaped up/down cells (§17.11) are deliberately NOT applied
    # here at build time -- §17.13's correction, same reasoning as Tasks'
    # own fix: applied dynamically instead, per row, the moment a row
    # actually becomes a template -- see modRecurringTasks.bas's
    # RefreshButtonStyling.
    # Centered ID, Total Subtasks, Subtask of, Deadline Offset (§17.11/
    # §17.14, same fields as Tasks' own fix, adjusted for the columns
    # this sheet actually has -- no Time Spent here, but Deadline Offset
    # is Recurring Tasks' own numeric-column equivalent).
    center = Alignment(horizontal="center", vertical="center")
    for col_letter in (id_cols["id"], task_cols["num_subtasks"], task_cols["subtask_of"], deadline_offset_col):
        for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
            ws[f"{col_letter}{row}"].alignment = center

    # Sequential blue ramp on Number of Subtasks per §17.5, same as Tasks
    # (PLAN.md §19 Phase 6).
    cf.apply_rules(
        ws,
        cf.sequential_rules(
            task_cols["num_subtasks"],
            FIRST_DATA_ROW,
            LAST_DATA_ROW,
            buckets=[(1, 100), (3, 150), (6, 200), (11, 250), (21, 300)],
        ),
    )

    # --- Sheet tab color (reference sheet, §17.9) ---
    ws.sheet_properties.tabColor = ds.TAB_REFERENCE

    return ws


def _apply_depth_shading(ws, level_col_letter):
    """Same treatment as Tasks §6.4 -- progressively lighter row shading
    per hierarchy depth, across the full visible column range (A:N here,
    vs. Tasks' A:P -- this sheet has 14 visible columns, not 16).
    """
    from openpyxl.formatting.rule import FormulaRule

    visible_range = f"A{FIRST_DATA_ROW}:N{LAST_DATA_ROW}"
    level_cell = f"${level_col_letter}{FIRST_DATA_ROW}"

    shading_steps = [100, 150, 200, 250, 300]
    for depth, step in enumerate(shading_steps, start=1):
        fill = PatternFill("solid", bgColor=ds.SEQUENTIAL_BLUE[step])
        rule = FormulaRule(formula=[f"{level_cell}={depth}"], fill=fill)
        ws.conditional_formatting.add(visible_range, rule)


def _apply_subtask_prefix(ws, level_col_letter, task_col_letter):
    """Same "⤷ " subtask prefix as Tasks (see tasks.py's own
    _apply_subtask_prefix docstring for the full rationale) -- applied
    here too for consistency, since Recurring Tasks shares the identical
    hierarchy engine and depth-shading mechanism.
    """
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.styles.numbers import NumberFormat

    task_range = f"{task_col_letter}{FIRST_DATA_ROW}:{task_col_letter}{LAST_DATA_ROW}"
    level_cell = f"${level_col_letter}{FIRST_DATA_ROW}"

    rule = Rule(type="expression", formula=[f"{level_cell}>=1"])
    # See tasks.py's own _apply_subtask_prefix for why numFmtId is
    # required; 202 here just to be distinct from Tasks' own 201.
    rule.dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=202, formatCode='"⤷ "@'))
    ws.conditional_formatting.add(task_range, rule)
