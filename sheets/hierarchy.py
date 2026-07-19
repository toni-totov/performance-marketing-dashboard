"""Reusable hierarchy-engine column scaffolding for openpyxl-built sheets.

Used by both Tasks (PLAN.md §6) and Recurring Tasks (§12.2) to lay out the
shared "self-organizing checklist" columns: visible ID / up-down reorder /
Task name / Total Subtasks / Subtask of, plus the hidden helper
columns (a stable key, parent key, depth level, and sibling-order
tiebreaker) that the VBA RenumberHierarchy() engine (see
vba/modHierarchy.bas) operates on at runtime.

Split into three independently-placeable pieces rather than one
contiguous block, because the columns *in between* differ per sheet:
Tasks needs `Start Date` sandwiched between the ID/reorder columns and
`Task` (§6.1), while Recurring Tasks prepends `Status` before everything
(§12.4) but otherwise keeps the block contiguous. Each sheet's own build
module calls these at whatever column positions its own layout needs, and
places its remaining columns (Priority, Deadline, Timer, etc. for Tasks;
Recurrence, Task Start On, etc. for Recurring Tasks) around them. Hidden
columns are conventionally grouped together at the end of a sheet's full
column list (after its last visible column), not immediately after
`Subtask of`.
"""

from openpyxl.utils import get_column_letter

from . import design_system as ds

# (logical name, header label) for the two column groups, in write order.
ID_REORDER_HEADERS = [
    ("id", "ID"),
    ("up", "↑"),
    ("down", "↓"),
]

TASK_HEADERS = [
    ("task", "Task"),
    ("num_subtasks", "Total Subtasks"),
    ("subtask_of", "Subtask of"),
]

# Logical names for the shared hidden columns, in write order. The first
# two get their header *text* from the caller (key_col_name /
# parent_key_col_name), since Tasks and Recurring Tasks name them
# differently for the same mechanism.
HIDDEN_KEYS = ["key", "parent_key", "level", "sibling_order"]


def _write_header_group(ws, start_col, headers, header_row):
    """Write `headers` (a list of (logical_name, label)) starting at
    `start_col`, formatted per the Visual Design System (PLAN.md §17:
    brand-accent header band, white bold text, wrapped + a right-edge
    divider per §17.11's post-review addendum). Returns a dict mapping
    logical name -> column letter. Caller is responsible for setting the
    header row's own height to `design_system.HEADER_ROW_HEIGHT`.
    """
    columns = {}
    col = start_col
    for key, label in headers:
        letter = get_column_letter(col)
        cell = ws.cell(row=header_row, column=col, value=label)
        ds.style_header_cell(cell)
        columns[key] = letter
        col += 1
    return columns


def write_id_reorder_columns(ws, start_col, header_row=1):
    """Write `ID`, `↑`, `↓` starting at `start_col`. Returns a dict
    mapping logical name ("id"/"up"/"down") -> column letter.
    """
    return _write_header_group(ws, start_col, ID_REORDER_HEADERS, header_row)


def write_task_columns(ws, start_col, header_row=1):
    """Write `Task`, `Total Subtasks`, `Subtask of` starting at
    `start_col`. Returns a dict mapping logical name ("task"/
    "num_subtasks"/"subtask_of") -> column letter.
    """
    return _write_header_group(ws, start_col, TASK_HEADERS, header_row)


def write_hidden_hierarchy_columns(ws, start_col, key_col_name, parent_key_col_name, header_row=1):
    """Write the 4 hidden hierarchy columns starting at `start_col`:
    the stable key, its parent-key, depth Level, and SiblingOrder.

    `key_col_name` / `parent_key_col_name` are the header text to use for
    the two key columns (e.g. "TaskKey"/"ParentKey" for Tasks,
    "RecurringKey"/"RecurringParentKey" for Recurring Tasks) -- the
    mechanism is identical, only the names differ per sheet.

    Returns a dict mapping logical name ("key"/"parent_key"/"level"/
    "sibling_order") -> column letter. Each written column is also marked
    hidden.
    """
    labels = [key_col_name, parent_key_col_name, "Level", "SiblingOrder"]
    headers = list(zip(HIDDEN_KEYS, labels))
    columns = _write_header_group(ws, start_col, headers, header_row)
    for letter in columns.values():
        ws.column_dimensions[letter].hidden = True
    return columns


def number_of_subtasks_formula(id_col_letter, row, data_first_row, data_last_row):
    """PLAN.md §6.2 point 3: live recursive descendant count via ID-prefix
    wildcard match against the sheet's own ID column -- no VBA, no
    recursion. `row` is the row this formula will live on; the ID range
    spans `data_first_row`..`data_last_row`. Blank instead of a literal
    "0" whenever the count is zero (an empty/unused row, or a genuine
    leaf task with no children) -- found live (PLAN.md §19 Phase 5) that
    a "0" on every never-used row read as visual noise. No VBA code reads
    this cell expecting a literal "0" text value specifically (every
    consumer already checks Len(...)>0 first, which blank satisfies the
    same way "0" did), so this is safe on both Tasks and Recurring Tasks.
    """
    id_range = f"${id_col_letter}${data_first_row}:${id_col_letter}${data_last_row}"
    this_id = f"{id_col_letter}{row}"
    count_expr = f'COUNTIF({id_range},{this_id}&".*")'
    return f'=IF({count_expr}=0,"",{count_expr})'


def next_sibling_order(existing_orders, step=10):
    """PLAN.md §6.2 point 6: appending a task at the end of its siblings
    assigns max(existing siblings' SiblingOrder) + step. Pass an empty
    sequence for the first child of a parent.
    """
    return (max(existing_orders) + step) if existing_orders else step
