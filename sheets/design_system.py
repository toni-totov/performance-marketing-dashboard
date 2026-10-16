"""Shared color, typography, and icon constants — the single source of
truth for PLAN.md §17 (Visual Design System).

Every sheet-building module imports from here rather than hardcoding hex
values, so the whole workbook stays visually consistent and PLAN.md stays
the one place these decisions live.

Colors are bare 6-digit hex strings (no leading "#"), ready to pass
directly to openpyxl's PatternFill(fgColor=...) / Font(color=...).
"""

# --- Brand / chrome accent (§17.1, §17.17) --------------------------------
# Structural elements only: header bands, primary buttons, sheet tabs,
# Dashboard card accent borders. Always paired with white text — black
# text on this fill fails contrast (2.4:1).
BRAND_ACCENT = "1F4E79"        # corporate navy (§17.17, was the earlier brand-blue candidate #1428A0)
BRAND_ACCENT_TEXT = "FFFFFF"

# --- Third-party platform brand colors (§17.10, PLAN.md §19 Phase 6) -----
# A second, narrower deliberate exception to "identity chips always draw
# from CATEGORICAL" (the same kind of exception BRAND_ACCENT itself
# already is) -- user-reported: Monthly Media Plan's Google/META chips and
# Gantt bars used arbitrary categorical slots instead of colors anyone
# actually recognizes as those two platforms. Real single-tone brand
# colors instead, scoped to exactly this one field -- not added to
# CATEGORICAL itself, since that set's own hues are chosen and ordered as
# a validated whole (see CATEGORICAL's own docstring) and platform
# identity isn't a "pick the next available slot" categorical case the
# way Funded/Optimization/Recurrence/Location are.
PLATFORM_GOOGLE = "4285F4"     # Google Blue
PLATFORM_META = "0866FF"       # Meta Blue

# --- Categorical data palette (§17.1, §17.2) ------------------------------
# Fixed 8-hue order, chosen and validated together as a set — never cycle,
# never swap a single slot. Used for identity chips (Funded, Platform,
# Optimization, Recurrence, Location, chart series). See §17.2 for the
# exact slot assignment per field.
CATEGORICAL = [
    "2A78D6",  # slot 1 — blue
    "1BAF7A",  # slot 2 — aqua
    "EDA100",  # slot 3 — yellow
    "008300",  # slot 4 — green
    "4A3AA7",  # slot 5 — violet
    "E34948",  # slot 6 — red
    "E87BA4",  # slot 7 — magenta
    "EB6834",  # slot 8 — orange
]

# --- Status / severity palette (§17.3) ------------------------------------
# Fixed, reserved meaning — never reused for series identity. Always paired
# with an icon/label, never color alone.
STATUS_GOOD = "0CA30C"
STATUS_WARNING = "FAB219"
STATUS_SERIOUS = "EC835A"
STATUS_CRITICAL = "D03B3B"
WARNING_SOFT = "FCE7B3"
GRIDLINE_DARK = "BDBBB5"       # faint placeholder text, e.g. Linked Campaign "None" (V4)        # pale amber -- "soon", one step below STATUS_WARNING (V4, PLAN.md §22.8)

# Priority (Tasks, Recurring Tasks): Low/Medium/High -> good/warning/critical
PRIORITY_LOW = STATUS_GOOD
PRIORITY_MEDIUM = STATUS_WARNING
PRIORITY_HIGH = STATUS_CRITICAL

# Urgency/deadline coloring (Days Left, Overdue, reminders):
URGENCY_OVERDUE = STATUS_CRITICAL
URGENCY_DUE_SOON = STATUS_WARNING
URGENCY_ON_TRACK = STATUS_GOOD

# --- Chrome & ink (§17.6) --------------------------------------------------
INK_PRIMARY = "0B0B0B"         # body text, values
INK_SECONDARY = "52514E"       # captions, helper text
INK_MUTED = "898781"           # column headers, axis-like labels
GRIDLINE = "E1E0D9"            # hairline border
PAGE_BG = "F9F9F7"             # subtle off-white sheet background
CARD_SURFACE = "FCFCFB"        # KPI/stat card surface

# --- Workflow-status mini-scheme (§17.4) -----------------------------------
# Reused across every lifecycle Status field (Tasks, Campaigns) plus binary
# Active/Inactive and Yes/No fields.
WORKFLOW_NOT_STARTED = INK_MUTED       # Tasks "Not Started", Campaigns "Planned"
WORKFLOW_IN_MOTION = CATEGORICAL[0]    # Tasks "In Progress", Campaigns "Active"
WORKFLOW_DONE = STATUS_GOOD            # Completed
WORKFLOW_SKIPPED = "D3D1CB"            # Tasks "Skipped" (V4, PLAN.md §22.4) -- lighter than Not Started: "won't happen"
WORKFLOW_RESCHEDULED = STATUS_SERIOUS  # Tasks "Rescheduled" (V4) -- orange: moved, still needs doing

WORKFLOW_ACTIVE = STATUS_GOOD          # Recurring Tasks / Schedule Meetings
WORKFLOW_INACTIVE = INK_MUTED

ACCESS_YES = STATUS_GOOD               # Login Information "Access"
ACCESS_NO = INK_MUTED                  # not critical red -- just informational

# --- Sequential ramp (§17.5) -----------------------------------------------
# Single blue hue, light -> dark = low -> high magnitude. Campaigns
# Progress %, Number of Subtasks data bars.
SEQUENTIAL_BLUE = {
    100: "CDE2FB",
    150: "B7D3F6",
    200: "9EC5F4",
    250: "86B6EF",
    300: "6DA7EC",
    350: "5598E7",
    400: "3987E5",
    450: "2A78D6",
    500: "256ABF",
    550: "1C5CAB",
    600: "184F95",
    650: "104281",
    700: "0D366B",
}

# --- Diverging pair (§17.5) -------------------------------------------------
# Archive's frozen Days Left ("days early/late at completion"): positive
# (early) -> blue arm, negative (late) -> red arm, gray neutral midpoint.
DIVERGING_POSITIVE = CATEGORICAL[0]    # blue -- finished early
DIVERGING_NEGATIVE = CATEGORICAL[5]    # red -- finished late
DIVERGING_MIDPOINT = "F0EFEC"          # gray -- exactly on time

# --- Typography (§17.7) -----------------------------------------------------
FONT_NAME = "Segoe UI"
# 16/24, not 15/22 (§17.16, user-requested bolder hierarchy) -- both stay
# within §17.7's own original range ("14-16pt" titles, "20-24pt" KPI
# numbers), just pushed to the top of each.
FONT_SIZE_TITLE = 16        # sheet/section titles
FONT_SIZE_KPI = 24          # Dashboard KPI big numbers
FONT_SIZE_HEADER = 10.5     # column headers (muted/secondary ink)
FONT_SIZE_BODY = 10.5       # body/data
FONT_SIZE_CAPTION = 9       # captions (secondary ink)

# --- Iconography (§17.8) ----------------------------------------------------
ICON_NOT_STARTED = "⬜"      # White square
ICON_IN_PROGRESS = "\U0001F535"  # Blue circle
ICON_COMPLETED = "✅"        # Check mark
ICON_TIMER_RUNNING = "▶"    # Play triangle
ICON_NOTES = "\U0001F4DD"        # Memo
ICON_UP = "↑"
ICON_DOWN = "↓"
ICON_NAV_PREV = "◀"
ICON_NAV_NEXT = "▶"         # Same glyph as ICON_TIMER_RUNNING -- never
                                  # on the same sheet (Timer = Tasks only,
                                  # nav = Calendar/Schedule/Monthly Media
                                  # Plan/Campaigns only), so no real collision.
ICON_REFRESH = "\U0001F504"
ICON_CONFIRM = "✓"
ICON_CANCEL = "✕"
ICON_EXCHANGE_RATE = "\U0001F4B1"

# --- Sheet-tab color-coding (§17.9) -----------------------------------------
TAB_CORE = BRAND_ACCENT                # Dashboard, Tasks, Calendar, Campaigns, Schedule
TAB_ARCHIVE = INK_MUTED                # Archive, Campaigns Archive
TAB_REFERENCE = CATEGORICAL[1]         # Monthly Media Plan, Recurring Tasks,
                                        # Login Information, Quick Links, Instructions

# --- Modernization pass (§17.10, PLAN.md §19 Phase 6, user-requested) -------
# The workbook read as "a 2008 spreadsheet" -- traced to Excel's own stock
# chrome (row/col headers, 100% zoom, ~15pt default row height) plus every
# ListObject Table using the built-in TableStyleMedium2 style, whose own
# blue/white banding visually fights the workbook's custom brand navy
# header band. Fixed with app-like window chrome plus a quiet, hand-rolled
# table treatment instead of a built-in style -- see apply_modern_chrome()
# and conditional_formatting.row_banding_rule() / table_hairline_borders().
DEFAULT_ZOOM = 115
DEFAULT_ROW_HEIGHT = 19
TABLE_STRIPE = PAGE_BG                 # subtle banding fill, reuses the page-bg token

# --- Header band, round 2 (§17.11 addendum, PLAN.md §19 Phase 6,
# user-reported after reviewing the fixed file) ---------------------------
# Two related complaints: a long header label (e.g. "Number of Subtasks")
# silently clips against its neighbor in a narrow column instead of
# wrapping, and the header band reads as one undifferentiated blue bar
# with no visual break between columns. HEADER_ROW_HEIGHT gives a wrapped
# 2-line label room to actually show both lines; HEADER_DIVIDER is a
# muted-gray hairline between header cells -- distinct enough to read as
# a column boundary without competing with the brand-blue fill the way a
# brighter or white divider would.
HEADER_ROW_HEIGHT = 32
HEADER_DIVIDER = INK_MUTED


def style_header_cell(cell, wrap=True, divider=True):
    """The shared header-cell look (§17.9's header-row convention,
    extended per the above): brand navy fill, white bold text,
    centered, and now wrapped by default so a long label breaks onto a
    second line instead of clipping -- plus a right-edge divider so
    adjacent header cells read as distinct columns. Callers still need to
    set the header ROW's own height to `HEADER_ROW_HEIGHT` separately
    (a cell-level call can't reach its row), and the rightmost header
    cell in a block should pass `divider=False` -- a trailing divider
    with nothing to its right just reads as a stray line.
    """
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    cell.font = Font(name=FONT_NAME, size=FONT_SIZE_HEADER, bold=True, color=BRAND_ACCENT_TEXT)
    cell.fill = PatternFill("solid", fgColor=BRAND_ACCENT)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    if divider:
        cell.border = Border(right=Side(style="hair", color=HEADER_DIVIDER))


def apply_button_look(ws, column_letter, first_row, last_row, fill=None, border_color=None):
    """Gives a column of single-cell "buttons" (Tasks/Recurring Tasks'
    Timer cell, every sheet's up/down reorder cells) a visible clickable
    shape -- a card-surface fill plus a full hairline border -- instead
    of bare icon text sitting directly on the sheet background with no
    affordance that it's interactive (§17.11, user-reported: "so that we
    know that it is a button"). Applied once across the whole data range
    at build time, not tied to any particular row's current icon/text
    value -- VBA only ever rewrites these cells' *values* at runtime
    (`ICON_UP`/`ICON_DOWN`/timer text), never their fill or border, and
    `RenumberHierarchy` rewrites values back into fixed row positions
    rather than moving styled Cell objects around, so a fixed per-column
    style baked in here at build time survives any later reordering.
    """
    from openpyxl.styles import Alignment, Border, PatternFill, Side

    cell_fill = PatternFill("solid", fgColor=fill or CARD_SURFACE)
    side = Side(style="thin", color=border_color or INK_MUTED)
    border = Border(top=side, bottom=side, left=side, right=side)
    center = Alignment(horizontal="center", vertical="center")
    for row in range(first_row, last_row + 1):
        cell = ws[f"{column_letter}{row}"]
        cell.fill = cell_fill
        cell.border = border
        cell.alignment = center


def apply_modern_chrome(ws, zoom=DEFAULT_ZOOM, gridlines=False, row_height=DEFAULT_ROW_HEIGHT):
    """Applies the workbook-wide "app, not spreadsheet" window chrome to a
    single sheet: gridlines off, a slightly larger default zoom, and a
    bit more breathing room on row height. Called once per sheet's
    build(), replacing what used to be a bare `ws.sheet_view.
    showGridLines = False` line.

    **Correction (§17.10's original modernization pass also hid Excel's
    own row/column letter headers -- reverted, user-reported)**: without
    visible row numbers, the user couldn't tell which row a new task
    needed to go into, or navigate/reference cells by position at all --
    a real usability regression for a sheet the user actively types into,
    not just a style mismatch the way the built-in table banding was.
    Gridlines-off doesn't have this problem (it's purely decorative, no
    navigational information lost), so that stays; only the row/column
    header visibility was reverted.

    `defaultRowHeight` silently has no effect unless `customHeight` is
    also set on the same `sheetFormatPr` -- confirmed empirically (a row
    height of 15pt, Excel's own default, came back even after setting
    `defaultRowHeight = 19` alone) before rolling this out project-wide.
    """
    ws.sheet_view.showGridLines = gridlines
    ws.sheet_view.zoomScale = zoom
    ws.sheet_format.defaultRowHeight = row_height
    ws.sheet_format.customHeight = True


def apply_box_outline(ws, cell_range, style="thin", color=None):
    """Draws a complete rectangular outline around `cell_range` (an
    "A1:C3"-style string), correctly for BOTH a merged "button" cell and
    a genuine multi-cell card area -- PLAN.md §19 Phase 6 bugfix.

    **Root cause this fixes**: setting `.border` on only a merged range's
    top-left anchor cell (the pattern used project-wide for every bordered
    button -- Calendar's Cancel, Schedule's Reminder Settings, Monthly
    Media Plan's Set Exchange Rate/Start New Month, Dashboard's Refresh
    Now/KPI cards) renders an INCOMPLETE box in real Excel: a merged
    cell's outer border is drawn per-EDGE from whichever physical cell
    sits at that edge (the rightmost cell contributes the right border,
    the bottom cell the bottom border, etc.), not copied from the anchor
    cell the way fill/font are. Only the anchor cell's own top+left edges
    would show, leaving the button looking like it's sitting "outside its
    box" on the right/bottom -- exactly what the user reported live.
    Confirmed via COM against a rebuilt file.

    Fix: apply a border to every physical cell in the range, not just the
    anchor -- but only on the sides that sit on the range's own outer
    edge (a cell only gets `top` if it's in the top row, `left` if it's
    in the leftmost column, etc.), not all 4 sides unconditionally. A
    single-row or single-column range (any genuine Excel merge always is)
    collapses every cell's top row and bottom row into the same row, so
    each cell there correctly gets both, which is exactly the "full
    4-side border on every cell" fix a merge needs. A true multi-row,
    multi-column region (e.g. a non-merged card background) instead gets
    a clean outer box with no interior grid lines, since interior cells
    get no border sides at all -- unconditionally bordering every side of
    every cell would draw a visible grid there instead of a single outline.
    """
    from openpyxl.styles import Border, Side
    from openpyxl.utils import range_boundaries

    side = Side(style=style, color=color or GRIDLINE)
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            ws.cell(row=row, column=col).border = Border(
                top=side if row == min_row else None,
                bottom=side if row == max_row else None,
                left=side if col == min_col else None,
                right=side if col == max_col else None,
            )


def modern_table_style():
    """A ListObject TableStyleInfo with no built-in named style at all
    (`name=None`) -- Excel renders the table with zero default styling
    (no banding, no borders), leaving every visual decision to our own
    header formatting + `conditional_formatting.row_banding_rule()` +
    `table_hairline_borders()`. Verified empirically: opens cleanly with
    no repair prompt, `ListObject.TableStyle` reads back blank, distinct
    from the previous per-sheet `TableStyleInfo(name="TableStyleMedium2",
    showRowStripes=True)` (a built-in style whose own blue/white banding
    was fighting the custom brand navy header band -- the main driver
    of the "dated spreadsheet" look this modernization pass addresses).
    """
    from openpyxl.worksheet.table import TableStyleInfo

    return TableStyleInfo(
        name=None, showFirstColumn=False, showLastColumn=False, showRowStripes=False, showColumnStripes=False
    )


def apply_table_hairline_borders(ws, first_col, last_col, first_row, last_row):
    """Direct (non-CF) hairline bottom border in `GRIDLINE`, applied to
    every cell in a table body -- the other half of the modern-table
    treatment alongside `modern_table_style()` (no built-in style) and
    `conditional_formatting.row_banding_rule()` (banding). No vertical
    borders and no top border -- a single bottom rule per row already
    reads as row separation without the boxed-in "classic Excel grid"
    look; letting the banding do the rest of the visual work.
    """
    from openpyxl.styles import Border, Side
    from openpyxl.utils import get_column_letter

    border = Border(bottom=Side(style="hair", color=GRIDLINE))
    for row in range(first_row, last_row + 1):
        for col in range(first_col, last_col + 1):
            ws.cell(row=row, column=col).border = border


def notes_icon_format():
    """§6.3: the shared Notes-column number format. Masks a cell's real
    text behind a fixed icon whenever non-empty (blank cells naturally
    show nothing, no format needed for that case). 4-section format --
    the last section is what applies to TEXT values. Shared by every
    sheet with a Notes column (Tasks, Archive, Campaigns, Campaigns
    Archive, Recurring Tasks) so the format string is defined once, not
    re-typed per sheet.
    """
    return f'General;General;General;"{ICON_NOTES}"'
