"""Calendar handoff contract (PLAN.md §8.1) -- a generic 4-value handoff
that lets any sheet launch the Calendar date-picker and get a value
written back to a specific cell, without Calendar needing to know
anything about the origin sheet ahead of time.

This module reserves the 4 handoff cells in a utility column on the
Calendar sheet (during Calendar's own build, Phase 2) and registers each
as a workbook-scoped defined name, so VBA on any origin sheet (e.g.
Tasks' Deadline double-click handler) can set them by name --
`HandoffOriginSheet`, `HandoffKey`, `HandoffKeyColumn`,
`HandoffTargetColumn` -- before navigating to Calendar, without needing to
know Calendar's exact layout. Calendar's Confirm button reads the same
names to know where to write the picked value back to.
"""

from openpyxl.workbook.defined_name import DefinedName

HANDOFF_NAMES = [
    "HandoffOriginSheet",
    "HandoffKey",
    "HandoffKeyColumn",
    "HandoffTargetColumn",
]


def write_handoff_cells(wb, calendar_sheet_name, anchor_row=1, anchor_col_letter="Z"):
    """Reserve 4 cells in a utility column on the Calendar sheet (default
    column Z, clear of the visible grid) for the handoff values, and
    register each as a workbook-scoped defined name.

    Returns a dict mapping handoff name -> its full sheet-qualified
    address, mainly for the caller's own bookkeeping/testing.
    """
    ws = wb[calendar_sheet_name]
    addresses = {}
    for i, name in enumerate(HANDOFF_NAMES):
        row = anchor_row + i
        ws.cell(row=row, column=_col_letter_to_index(anchor_col_letter))
        cell_ref = f"'{calendar_sheet_name}'!${anchor_col_letter}${row}"
        wb.defined_names[name] = DefinedName(name, attr_text=cell_ref)
        addresses[name] = cell_ref
    return addresses


def _col_letter_to_index(letter):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)
