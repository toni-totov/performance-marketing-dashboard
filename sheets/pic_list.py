"""PIC List -- the manager variant's maintained roster of team-member
names, feeding Tasks' PIC dropdown (PLAN.md manager-variant design). Manager-
only; not built for the employee variant at all, since employee Tasks has
no PIC column to source from it.

Deliberately a plain, contiguous, single-column growable list (not
Campaigns' fixed-slot layout) -- a person roster has no other structured
data tied to it the way a Campaign row does, so there's no need for
Campaigns' non-contiguous-slot machinery or its VBA-maintained mirror
column. What IS reused from that precedent is the *reason* CampaignNameList
exists in the first place: a plain fixed-size Data Validation list range
with merely-cleared cells is a documented, already-live-confirmed bug in
this codebase (a trailing selectable blank row in the actual dropdown UI).
The `PICList` Defined Name below uses the same dynamic OFFSET+COUNTA
technique for that reason, with one addition -- `MAX(1, COUNTA(...))` --
because this list starts genuinely empty (unlike Campaigns, which always
has a permanent "None" floor row): a bare COUNTA of zero would produce a
zero-height OFFSET, which Data Validation can show as an outright
evaluation error the moment the dropdown is opened before any name has
been added. The MAX floor keeps the range at a harmless height-1 reference
into a still-blank cell until real entries exist, then COUNTA alone takes
over correctly (no trailing blank) once there's at least one real name --
no permanent placeholder row needed.

Recompaction (closing a gap left by clearing a name mid-list, so COUNTA
never undercounts past a gap) is handled by VBA (`PIC List.bas`'s
Worksheet_Change handler), not here -- this module only builds structure.

Up/Down reorder columns (B/C, user-requested after using the manager
build) mirror Tasks' own identifier-then-buttons layout (ID, Up, Down,
...) -- PIC Name stands in for Tasks' ID as the "identifier" column here.
Deliberately NOT styled as clickable buttons at build time across all 200
allowance rows -- `design_system.apply_button_look` exists for that but
is unused project-wide; every sheet with Up/Down (Tasks, Recurring Tasks)
learned the hard way (PLAN.md §19 Phase 6 / §17.13 correction) that baking
button borders into a large pre-formatted allowance makes empty rows look
clickable, so styling is applied dynamically per real row by VBA instead
(`modTasksManager.bas`'s `RefreshPICListButtonStyling`).

Color column (D, user-requested: "would it be better visually to have
each PIC with different color"). Deliberately a plain empty cell the
manager colors herself via Excel's own native Fill Color picker, not a
custom color-picker UI -- she already knows that control, and it avoids
building a UserForm just to reproduce something Excel already offers.
VBA (`modTasksManager.bas`'s `RefreshButtonStyling`/`HandleRefreshPICColorsClick`)
reads each swatch's `Interior.Color` and applies it to that PIC's cell
wherever the name appears in Tasks/Archive; a cell left "No Fill"
(`Interior.Pattern = xlNone`) means "no custom color for this person" and
is left alone downstream. **Important constraint driving the button
below:** Excel/VBA has no event that fires when a cell's FORMATTING (as
opposed to its value) changes -- `Worksheet_Change` never sees a Fill
Color edit -- so color changes can't propagate automatically the instant
she picks one. A "Refresh PIC Colors" banner (mirroring Tasks' own row-
extender banner's build/click pattern) gives her an explicit, immediate
way to push a color change out; `ThisWorkbookManager.bas`'s
`Workbook_Open` also runs it once as a backfill, same convention as
`RecompactPICList`/`RefreshPICListButtonStyling`.

**Deliberately kept OUTSIDE `PICListTable`'s own `ref` range** -- confirmed
live via COM that this is not cosmetic: an Excel Table auto-expands to
include a newly-typed row (e.g. typing a name into the row right below
the table), and when it does, it carries the ROW ABOVE's formatting into
every column the table spans, fill color included. With Color inside the
table, a genuinely brand-new PIC row silently inherited whichever color
happened to sit in the row above it the moment the table grew to include
it -- confirmed reproducible via a plain sequence of cell writes, no VBA
even involved, so this isn't a scenario she could avoid by using the UI
carefully. Name/Up/Down don't show this symptom only because
`RefreshPICListButtonStyling` already unconditionally overwrites Up/Down
on every edit, masking whatever got auto-copied there; Color has no such
per-edit overwrite (it's manager-chosen persistent data, not something
VBA can reset to a known default). Keeping Color structurally outside the
ListObject means the table's own growth never touches column D at all --
confirmed via the same COM repro that this removes the leak entirely.

**Row banding stops at C, deliberately excluding Color** -- a second,
separate bug (user-reported: "row 1 I can't color, row 2 colors, row 3
doesn't..."). The row-banding conditional-format fill (see build()'s own
call to `cf.row_banding_rule`) alternates every other row by design, and
a CF fill visually overrides a cell's own manually-set direct fill
wherever the CF's condition is true -- so a swatch on a "banded" row
looked uncolored no matter what she picked, while the very next row (not
banded) looked fine, exactly the alternating pattern reported. A swatch
cell's only job is to show her chosen color unconditionally, which no CF
fill sitting on top of it can be allowed to interfere with, so banding
was scoped to stop at column C.
"""

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table

from . import conditional_formatting as cf
from . import design_system as ds

SHEET_NAME = "PIC List"
HEADER_ROW = 1
FIRST_DATA_ROW = 2
LAST_ROW = FIRST_DATA_ROW + 200 - 1  # 201 -- a person roster needs no row-extender
REFRESH_BUTTON_ROW = LAST_ROW + 1  # 202 -- fixed, unlike Tasks' banner, since this list never grows

HEADERS = ["PIC Name", ds.ICON_UP, ds.ICON_DOWN, "Color"]
_COL = {name: i + 1 for i, name in enumerate(HEADERS)}


def build(wb):
    """Create the PIC List sheet in `wb` (an openpyxl Workbook). Manager
    variant only. Returns the created worksheet.
    """
    ws = wb.create_sheet(SHEET_NAME)
    ds.apply_modern_chrome(ws)
    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=1)
    ws.row_dimensions[HEADER_ROW].height = ds.HEADER_ROW_HEIGHT

    for col, label in enumerate(HEADERS, start=1):
        cell = ws.cell(row=HEADER_ROW, column=col, value=label)
        ds.style_header_cell(cell)

    ws.column_dimensions[get_column_letter(_COL["PIC Name"])].width = 28
    ws.column_dimensions[get_column_letter(_COL[ds.ICON_UP])].width = 4
    ws.column_dimensions[get_column_letter(_COL[ds.ICON_DOWN])].width = 4
    ws.column_dimensions[get_column_letter(_COL["Color"])].width = 10

    # PICListTable spans only Name/Up/Down (A:C) -- Color (D) is
    # deliberately NOT part of the ListObject at all, see module
    # docstring's own correction note on why.
    table_last_col_letter = get_column_letter(len(HEADERS) - 1)
    table = Table(displayName="PICListTable", ref=f"A{HEADER_ROW}:{table_last_col_letter}{FIRST_DATA_ROW}")
    table.tableStyleInfo = ds.modern_table_style()
    ws.add_table(table)

    # Borders cover Color too, for visual consistency with the rest of
    # the row -- borders don't conflict with anything she sets on the
    # swatch. Row BANDING deliberately stops at C (does NOT cover Color)
    # -- confirmed live (user-reported): a conditional-format fill visibly
    # overrides a manually-set direct fill wherever the CF condition is
    # true, and row banding's condition alternates every other row --
    # exactly the "row 1 colors, row 2 doesn't, row 3 does..." pattern
    # she saw. A swatch cell's whole purpose is to show HER chosen color
    # unconditionally, which is fundamentally incompatible with any CF
    # fill sitting on top of it, banding included.
    last_col_letter = get_column_letter(len(HEADERS))
    banding_last_col_letter = get_column_letter(len(HEADERS) - 1)
    cf.apply_rules(ws, cf.row_banding_rule(f"A{FIRST_DATA_ROW}:{banding_last_col_letter}{LAST_ROW}", FIRST_DATA_ROW))
    ds.apply_table_hairline_borders(ws, 1, len(HEADERS), FIRST_DATA_ROW, LAST_ROW)

    name_col_letter = get_column_letter(_COL["PIC Name"])
    name_range = f"'{SHEET_NAME}'!${name_col_letter}${FIRST_DATA_ROW}:${name_col_letter}${LAST_ROW}"
    wb.defined_names["PICList"] = DefinedName(
        "PICList",
        attr_text=(
            f"OFFSET('{SHEET_NAME}'!${name_col_letter}${FIRST_DATA_ROW},0,0,"
            f"MAX(1,COUNTA({name_range})),1)"
        ),
    )

    # --- "Refresh PIC Colors" banner -- same build pattern as Tasks' own
    # row-extender banner (sheets/tasks.py), but at a fixed row (this list
    # never grows, so there's no moving-target position to track).
    ws.merge_cells(f"A{REFRESH_BUTTON_ROW}:D{REFRESH_BUTTON_ROW}")
    banner = ws.cell(row=REFRESH_BUTTON_ROW, column=1, value=f"{ds.ICON_REFRESH} Refresh PIC Colors")
    banner.font = Font(name=ds.FONT_NAME, size=ds.FONT_SIZE_BODY, bold=True, color=ds.INK_PRIMARY)
    banner.fill = PatternFill("solid", fgColor=ds.CARD_SURFACE)
    ds.apply_box_outline(ws, f"A{REFRESH_BUTTON_ROW}:D{REFRESH_BUTTON_ROW}", color=ds.INK_MUTED)
    banner.alignment = Alignment(horizontal="center", vertical="center")

    ws.sheet_properties.tabColor = ds.TAB_REFERENCE

    return ws
