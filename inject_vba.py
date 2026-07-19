"""Phase 4: injects all VBA (standard modules, sheet code-behind,
UserForms, ThisWorkbook) into the openpyxl-generated .xlsx via pywin32
COM automation, producing the final macro-enabled .xlsm.

Requires "Trust access to the VBA project object model" enabled in
Excel's Trust Center (confirmed done, Phase 0) -- otherwise VBProject
access raises a COM error.

Run build_dashboard.py first to produce the input .xlsx; this script
does not rebuild it, to keep structure generation and VBA injection as
separate concerns (PLAN.md §19 Phase 4).

See PLAN.md §18 for why UserForms are built from Python specs via VBIDE
automation rather than imported as .frm files, and for the CodeName/
ThisWorkbook locale-lookup gotchas this script works around -- both
confirmed live during Phase 3's own testing on this Bulgarian-locale
build machine, not theoretical concerns.
"""

import os
import sys

import win32com.client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sheets.forms._builder import build_form_from_spec
from sheets.forms import (
    exchange_rate_form,
    meeting_reminder_form,
    notes_form,
    reminder_settings_form,
    schedule_day_form,
)

VBA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vba")

VBEXT_CT_STDMODULE = 1
VBEXT_CT_DOCUMENT = 100
XL_OPENXML_WORKBOOK_MACRO_ENABLED = 52

# Standard modules. Injection order doesn't affect compilation (VBA
# resolves Public symbols across modules regardless of add order), kept
# in a stable, readable order here.
STANDARD_MODULES = [
    "modDesignSystem",
    "modHierarchy",
    "modCalendarHandoff",
    "modTasks",
    "modCampaigns",
    "modRecurringTasks",
    "modCalendar",
    "modMonthlyMediaPlan",
    "modSchedule",
]

# (sheet tab name, .bas file name) -- injected directly into that sheet's
# own CodeModule (PLAN.md §18's sheet-named-.bas convention).
SHEET_CODE_BEHIND = [
    ("Tasks", "Tasks.bas"),
    ("Archive", "Archive.bas"),
    ("Calendar", "Calendar.bas"),
    ("Campaigns", "Campaigns.bas"),
    ("Campaigns Archive", "Campaigns Archive.bas"),
    ("Recurring Tasks", "Recurring Tasks.bas"),
    ("Monthly Media Plan", "Monthly Media Plan.bas"),
    ("Schedule Data", "Schedule Data.bas"),
    ("Schedule", "Schedule.bas"),
    ("Dashboard", "Dashboard.bas"),
    ("Login Information", "Login Information.bas"),
]

FORM_SPECS = [
    notes_form.SPEC,
    exchange_rate_form.SPEC,
    schedule_day_form.SPEC,
    reminder_settings_form.SPEC,
    meeting_reminder_form.SPEC,
]


def _strip_attribute_line(text):
    lines = text.splitlines()
    return "\n".join(ln for ln in lines if not ln.strip().startswith("Attribute VB_Name"))


def _read_bas(filename):
    path = os.path.join(VBA_DIR, filename)
    with open(path, "r", encoding="utf-8") as f:
        return _strip_attribute_line(f.read())


def _find_thisworkbook_component(wb):
    """The ThisWorkbook component's own VBA name is locale-dependent
    (confirmed on this Bulgarian-locale build machine -- a fresh, never-
    previously-VBA-touched workbook does NOT auto-assign it literally
    "ThisWorkbook"), and unlike a sheet it has no tab name to look up by.
    Found by elimination instead: the one Type=vbext_ct_Document (100)
    component whose Name isn't any sheet's own CodeName. Must run after
    at least one VBComponents.Add call -- Worksheet.CodeName returns ""
    until the VBA project has been touched by one (PLAN.md §18).
    """
    sheet_codenames = set()
    for i in range(1, wb.Sheets.Count + 1):
        sheet_codenames.add(wb.Sheets(i).CodeName)
    for comp in wb.VBProject.VBComponents:
        if comp.Type == VBEXT_CT_DOCUMENT and comp.Name not in sheet_codenames:
            return comp
    raise RuntimeError("Could not find the ThisWorkbook component")


def inject(input_xlsx, output_xlsm):
    """Opens input_xlsx, injects all VBA, saves as a genuine macro-
    enabled output_xlsm (a plain .xlsx save would silently strip the VBA
    project). Raises on any failure rather than producing a silently
    partial result.
    """
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(input_xlsx)

        for modname in STANDARD_MODULES:
            comp = wb.VBProject.VBComponents.Add(VBEXT_CT_STDMODULE)
            comp.Name = modname
            comp.CodeModule.AddFromString(_read_bas(modname + ".bas"))
            print(f"Injected standard module: {modname}")

        for spec in FORM_SPECS:
            build_form_from_spec(wb, spec)
            print(f"Injected UserForm: {spec['name']}")

        for sheet_name, bas_name in SHEET_CODE_BEHIND:
            code_name = wb.Sheets(sheet_name).CodeName
            comp = wb.VBProject.VBComponents(code_name)
            comp.CodeModule.AddFromString(_read_bas(bas_name))
            print(f"Injected {bas_name} into the {sheet_name} sheet module")

        thisworkbook_comp = _find_thisworkbook_component(wb)
        thisworkbook_comp.CodeModule.AddFromString(_read_bas("ThisWorkbook.bas"))
        print("Injected ThisWorkbook.bas into the ThisWorkbook module")

        wb.SaveAs(output_xlsm, FileFormat=XL_OPENXML_WORKBOOK_MACRO_ENABLED)
        print("Saved:", output_xlsm)

        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    from build_dashboard import OUTPUT_XLSX

    output_xlsm = os.path.splitext(OUTPUT_XLSX)[0] + ".xlsm"
    if not os.path.exists(OUTPUT_XLSX):
        raise SystemExit(f"{OUTPUT_XLSX} not found -- run build_dashboard.py first.")
    inject(OUTPUT_XLSX, output_xlsm)
