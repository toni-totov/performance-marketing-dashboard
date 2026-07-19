"""Schedule Data -- raw storage backing the visual Schedule calendar
(PLAN.md §13.1). All editing happens via the frmScheduleDay popup
(§13.3, Phase 3), never direct typing -- but visible (not Very Hidden
like Time Log), confirmed with the user, in case manual inspection or a
bulk fix is ever needed.

Two independent genuine auto-expanding Excel Tables side by side:
  Work Location (A:B): Date, Location -- one row per logged date; Date
    behaves as a de-facto unique key (the popup's Save upserts by date,
    §13.3), enforced by VBA, not a formula constraint.
  Meetings (D:J, leaving C as a gap column): Meeting Name, Start Time,
    End Time, Recurrence, Anchor Date, Status, hidden MeetingKey. A
    recurring meeting stays a single rule row (§13.4) -- never physically
    duplicated per occurrence, unlike Recurring Tasks' generated Tasks
    rows -- so this table only ever holds as many rows as there are
    meeting *rules*, not occurrences.

Both start at the minimum valid size (header + one data row) and grow via
VBA's ListRows.Add (Phase 3), same pattern as Time Log/Archive/Campaigns
Archive. Number formats live only on the initial data row (verified
inheriting correctly via ListRows.Add for those sheets); conditional
formatting and data validation are pre-applied across a generous fixed
row allowance instead (confirmed NOT to travel with Table growth,
Archive's own established finding) -- Work Location: 400 rows (~13
months of daily entries), Meetings: 60 rows (a rule list, not
per-occurrence, so a much shorter allowance fits -- confirmed with the
user for both numbers).
"""

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

from . import conditional_formatting as cf
from . import design_system as ds

SHEET_NAME = "Schedule Data"
HEADER_ROW = 1
FIRST_DATA_ROW = 2

_WORK_LOCATION_LAST_ROW = FIRST_DATA_ROW + 400 - 1  # 401
_MEETINGS_LAST_ROW = FIRST_DATA_ROW + 60 - 1  # 61

_DATE_FORMAT = "DD.MM.YYYY"
_TIME_FORMAT = "HH:MM"

_WORK_LOCATION_HEADERS = ["Date", "Location"]
_WL_COL = {name: i + 1 for i, name in enumerate(_WORK_LOCATION_HEADERS)}  # A=1, B=2

_MEETINGS_FIRST_COL = 4  # D, leaving C as a gap column
_MEETINGS_HEADERS = ["Meeting Name", "Start Time", "End Time", "Recurrence", "Anchor Date", "Status"]
_MT_COL = {name: _MEETINGS_FIRST_COL + i for i, name in enumerate(_MEETINGS_HEADERS)}
_MEETING_KEY_COL = _MEETINGS_FIRST_COL + len(_MEETINGS_HEADERS)  # J

_LOCATION_PAIRS = [
    ("Home Office", ds.CATEGORICAL[0]),
    ("Office", ds.CATEGORICAL[1]),
    ("Holiday", ds.CATEGORICAL[2]),
    # User-requested additions (PLAN.md §19 Phase 5): two more "away"
    # states alongside Holiday, distinct from each other so a day's
    # actual reason for being excluded stays visible later, not just
    # that it was excluded. All three (Holiday/Sick/Out of Office) --
    # not the two working states -- exclude that date from recurring
    # meeting occurrences (modSchedule.bas's MeetingOccursOn / this
    # sheet's badge formula both check for all three).
    ("Sick", ds.CATEGORICAL[3]),
    ("Out of Office", ds.CATEGORICAL[4]),
]
_RECURRENCE_PAIRS = [
    ("Daily", ds.CATEGORICAL[0]),
    # User-requested addition (PLAN.md §19 Phase 5): "Daily" always meant
    # literally every day, including weekends, with no way to restrict a
    # recurring meeting to working days only. "Weekdays" is a new,
    # separate recurrence type (not a modifier on "Daily") so existing
    # "Daily" meetings keep their exact current every-day meaning
    # unchanged. Slot 4 -- a fresh categorical color, not reusing/
    # reassigning Weekly/Monthly/One-time's existing slots 1-3.
    ("Weekdays", ds.CATEGORICAL[4]),
    ("Weekly", ds.CATEGORICAL[1]),
    ("Monthly", ds.CATEGORICAL[2]),
    ("One-time", ds.CATEGORICAL[3]),
]


def _style_header_cell(cell):
    ds.style_header_cell(cell)


def build(wb):
    """Create the Schedule Data sheet in `wb` (an openpyxl Workbook).
    Returns the created worksheet.
    """
    ws = wb.create_sheet(SHEET_NAME)
    ds.apply_modern_chrome(ws)
    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=1)
    ws.row_dimensions[HEADER_ROW].height = ds.HEADER_ROW_HEIGHT

    # --- Work Location table (A:B) ---
    for name, col in _WL_COL.items():
        _style_header_cell(ws.cell(row=HEADER_ROW, column=col, value=name))
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14

    ws.cell(row=FIRST_DATA_ROW, column=_WL_COL["Date"]).number_format = _DATE_FORMAT

    location_dv = DataValidation(type="list", formula1='"Home Office,Office,Holiday"', allow_blank=True)
    location_dv.add(f"B{FIRST_DATA_ROW}:B{_WORK_LOCATION_LAST_ROW}")
    ws.add_data_validation(location_dv)

    cf.apply_rules(ws, cf.categorical_rules("B", FIRST_DATA_ROW, _WORK_LOCATION_LAST_ROW, _LOCATION_PAIRS))

    wl_table = Table(displayName="WorkLocationTable", ref=f"A{HEADER_ROW}:B{FIRST_DATA_ROW}")
    wl_table.tableStyleInfo = ds.modern_table_style()
    ws.add_table(wl_table)

    cf.apply_rules(ws, cf.row_banding_rule(f"A{FIRST_DATA_ROW}:B{_WORK_LOCATION_LAST_ROW}", FIRST_DATA_ROW))
    ds.apply_table_hairline_borders(ws, 1, 2, FIRST_DATA_ROW, _WORK_LOCATION_LAST_ROW)

    # --- Meetings table (D:J, C left as a gap column) ---
    for name, col in _MT_COL.items():
        _style_header_cell(ws.cell(row=HEADER_ROW, column=col, value=name))
    _style_header_cell(ws.cell(row=HEADER_ROW, column=_MEETING_KEY_COL, value="MeetingKey"))

    widths = {
        _MT_COL["Meeting Name"]: 24,
        _MT_COL["Start Time"]: 12,
        _MT_COL["End Time"]: 12,
        _MT_COL["Recurrence"]: 12,
        _MT_COL["Anchor Date"]: 14,
        _MT_COL["Status"]: 12,
    }
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.column_dimensions[get_column_letter(_MEETING_KEY_COL)].hidden = True

    ws.cell(row=FIRST_DATA_ROW, column=_MT_COL["Start Time"]).number_format = _TIME_FORMAT
    ws.cell(row=FIRST_DATA_ROW, column=_MT_COL["End Time"]).number_format = _TIME_FORMAT
    ws.cell(row=FIRST_DATA_ROW, column=_MT_COL["Anchor Date"]).number_format = _DATE_FORMAT

    recurrence_letter = get_column_letter(_MT_COL["Recurrence"])
    status_letter = get_column_letter(_MT_COL["Status"])

    recurrence_dv = DataValidation(type="list", formula1='"One-time,Daily,Weekly,Monthly"', allow_blank=True)
    recurrence_dv.add(f"{recurrence_letter}{FIRST_DATA_ROW}:{recurrence_letter}{_MEETINGS_LAST_ROW}")
    ws.add_data_validation(recurrence_dv)

    status_dv = DataValidation(type="list", formula1='"Active,Inactive"', allow_blank=True)
    status_dv.add(f"{status_letter}{FIRST_DATA_ROW}:{status_letter}{_MEETINGS_LAST_ROW}")
    ws.add_data_validation(status_dv)

    cf.apply_rules(ws, cf.categorical_rules(recurrence_letter, FIRST_DATA_ROW, _MEETINGS_LAST_ROW, _RECURRENCE_PAIRS))
    cf.apply_rules(ws, cf.binary_state_rules(status_letter, FIRST_DATA_ROW, _MEETINGS_LAST_ROW, "Active", "Inactive"))

    meetings_last_col_letter = get_column_letter(_MEETING_KEY_COL)
    meetings_first_col_letter = get_column_letter(_MEETINGS_FIRST_COL)
    mt_table = Table(
        displayName="MeetingsTable",
        ref=f"{meetings_first_col_letter}{HEADER_ROW}:{meetings_last_col_letter}{FIRST_DATA_ROW}",
    )
    mt_table.tableStyleInfo = ds.modern_table_style()
    ws.add_table(mt_table)

    # Meetings table's own hairline/banding -- MeetingKey (hidden) column
    # included since it's still part of the ListObject's own range.
    cf.apply_rules(
        ws,
        cf.row_banding_rule(
            f"{meetings_first_col_letter}{FIRST_DATA_ROW}:{meetings_last_col_letter}{_MEETINGS_LAST_ROW}",
            FIRST_DATA_ROW,
        ),
    )
    ds.apply_table_hairline_borders(ws, _MEETINGS_FIRST_COL, _MEETING_KEY_COL, FIRST_DATA_ROW, _MEETINGS_LAST_ROW)

    # --- Sheet tab color: muted gray, same visual grouping as Archive
    # (§17.9 doesn't explicitly list Schedule Data, but it plays the same
    # "backing data store, not primary daily-use" role) ---
    ws.sheet_properties.tabColor = ds.TAB_ARCHIVE

    return ws
