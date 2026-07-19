"""Reusable Monday-start week/month grid formula generators.

Legacy-safe DATE/WEEKDAY arithmetic (no dynamic arrays, no VBA needed to
regenerate a grid) -- reused by:
  - Calendar (PLAN.md §8.2): the 6x7 month grid.
  - Monthly Media Plan (§9.3) and Campaigns (§10.3): Gantt week columns.
  - Schedule (§13.2): reuses Calendar's month grid directly.

All functions return formula strings (starting with "=") or formula
fragments meant to be embedded in a larger formula -- callers decide where
each lands on the sheet.
"""

# WEEKDAY's type-3 system: Monday=0 .. Sunday=6, which is what makes
# "date - WEEKDAY(date, 3)" land on that date's own Monday.
_WEEKDAY_MONDAY_START = 3

# Hardcoded English name arrays for locale-proof month/day-name lookups
# (PLAN.md §19 Phase 5, found live): TEXT()'s "mmm"/"mmmm"/"ddd"/"dddd"
# format codes render the month/day NAME in the workbook's system
# locale, with no format-code variant that forces a specific language --
# confirmed live on this Bulgarian-locale machine ("Jul" rendering as
# "юли"). This is a distinct bug from the already-documented TEXT()
# decimal-separator locale issue (CLAUDE.md) -- that one is about
# NUMBER formatting, this one is about NAME text. An INDEX() lookup
# against a hardcoded English array sidesteps Excel's locale handling
# entirely, since it never asks Excel to localize anything.
_MONTH_NAMES_EXCEL_ARRAY = '{"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"}'
_MONTH_NAMES_FULL_EXCEL_ARRAY = (
    '{"January","February","March","April","May","June","July",'
    '"August","September","October","November","December"}'
)
_DAY_NAMES_EXCEL_ARRAY = '{"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"}'


def month_name_formula(date_cell):
    """Locale-proof 3-letter month name (e.g. "Jul"), as a bare
    expression (no leading "=") meant to be embedded in a larger
    formula. See module-level comment above for why this exists instead
    of TEXT(date_cell,"mmm").
    """
    return f"INDEX({_MONTH_NAMES_EXCEL_ARRAY},MONTH({date_cell}))"


def month_name_full_formula(date_cell):
    """Locale-proof full month name (e.g. "July"), same reasoning as
    month_name_formula() -- for a TEXT(...,"mmmm") replacement.
    """
    return f"INDEX({_MONTH_NAMES_FULL_EXCEL_ARRAY},MONTH({date_cell}))"


def weekday_name_formula(date_cell):
    """Locale-proof full weekday name (e.g. "Monday"), as a bare
    expression. WEEKDAY(date,2) returns 1=Monday..7=Sunday, matching
    _DAY_NAMES_EXCEL_ARRAY's order (this module's Monday-start
    convention throughout).
    """
    return f"INDEX({_DAY_NAMES_EXCEL_ARRAY},WEEKDAY({date_cell},2))"


def week_number_formula(date_cell):
    """The calendar week number, as a bare expression. WEEKNUM's
    week-start argument (2 = Monday-start, matching this module's own
    convention) is a plain integer, never a localized string -- carries
    none of TEXT()'s month/day-name locale risk.
    """
    return f"WEEKNUM({date_cell},2)"


def month_grid_anchor_formula(year_cell, month_cell):
    """First grid cell: the Monday of the week containing day 1 of the
    given year/month. `year_cell` / `month_cell` are cell references
    (e.g. "SelectedYear", "SelectedMonth").
    """
    first_of_month = f"DATE({year_cell},{month_cell},1)"
    return f"={first_of_month}-WEEKDAY({first_of_month},{_WEEKDAY_MONDAY_START})"


def month_grid_formulas(anchor_cell):
    """The full 6-row x 7-column (Mon-Sun) month grid, row-major, as a
    list of lists of formula strings. Each cell is computed directly from
    `anchor_cell` (the first grid cell) rather than chained from its
    neighbor, so every formula is independently correct regardless of
    fill/copy order.
    """
    grid = []
    for row in range(6):
        row_formulas = []
        for col in range(7):
            offset = row * 7 + col
            formula = f"={anchor_cell}" if offset == 0 else f"={anchor_cell}+{offset}"
            row_formulas.append(formula)
        grid.append(row_formulas)
    return grid


def current_week_monday_formula():
    """TODAY()-relative Monday -- the rolling anchor for Campaigns' §10.3
    auto-advancing 15-week Gantt window.
    """
    return f"=TODAY()-WEEKDAY(TODAY(),{_WEEKDAY_MONDAY_START})"


def week_start_formula(anchor_cell, week_offset):
    """The Monday of the Nth week from `anchor_cell` (0 = the anchor
    week itself).
    """
    if week_offset == 0:
        return f"={anchor_cell}"
    return f"={anchor_cell}+{week_offset * 7}"


def week_end_formula(week_start_cell):
    """The Sunday of the week starting at `week_start_cell`."""
    return f"={week_start_cell}+6"


def week_label_formula(week_start_cell, week_end_cell):
    """A "06–12 Jul"-style label for a Gantt week-column header.
    Locale-proof (see module-level comment) -- uses month_name_formula()
    instead of TEXT(...,"mmm").
    """
    return (
        f'=TEXT({week_start_cell},"dd")&"–"&TEXT({week_end_cell},"dd")'
        f'&" "&{month_name_formula(week_end_cell)}'
    )


def gantt_bar_condition(start_date_cell, end_date_cell, week_start_cell, week_end_cell):
    """The overlap test that fills a Gantt week-cell: does
    [start_date, end_date] overlap [week_start, week_end] at all?
    Shared by Monthly Media Plan (§9.3) and Campaigns (§10.3) -- return
    value is a bare boolean expression (no leading "="), meant to be
    embedded in a conditional-formatting formula.
    """
    return f"AND({start_date_cell}<={week_end_cell},{end_date_cell}>={week_start_cell})"
