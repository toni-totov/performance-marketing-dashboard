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

import argparse
import os
import sys

import win32com.client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_config import output_xlsm_path, output_xlsx_path
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


def standard_modules(variant):
    """(component name, source .bas filename) pairs for every standard
    module, per variant (PLAN.md, "Manager-variant workbook" §9).
    Component name and filename differ for modRecurringTasksWriter --
    both variants inject a component literally named
    "modRecurringTasksWriter" (so modRecurringTasks.bas's own qualified
    calls into it resolve identically either way), but source the content
    from a different file -- same "one qualified name, variant-selected
    source file" mechanism already used for Tasks/Archive's sheet
    code-behind, extended here to a standard module (a gap found and
    fixed during this phase -- modRecurringTasks.bas originally called
    modTasks.* directly, which doesn't exist in the manager build; see
    modRecurringTasksWriter.bas's own header comment for the full
    rationale). Injection order doesn't affect compilation (VBA resolves
    Public symbols across modules regardless of add order), kept in a
    stable, readable order here.
    """
    is_manager = variant == "manager"
    shared = [
        ("modDesignSystem", "modDesignSystem.bas"),
        ("modHierarchy", "modHierarchy.bas"),
        ("modCalendarHandoff", "modCalendarHandoff.bas"),
        ("modCampaigns", "modCampaigns.bas"),
        ("modRecurringTasks", "modRecurringTasks.bas"),
        ("modCalendar", "modCalendar.bas"),
        ("modSchedule", "modSchedule.bas"),
    ]
    if is_manager:
        variant_specific = [
            ("modTasksManager", "modTasksManager.bas"),
            ("modRecurringTasksWriter", "modRecurringTasksWriterManager.bas"),
        ]
    else:
        variant_specific = [
            ("modTasks", "modTasks.bas"),
            ("modRecurringTasksWriter", "modRecurringTasksWriter.bas"),
            ("modMonthlyMediaPlan", "modMonthlyMediaPlan.bas"),
            ("modMigration", "modMigration.bas"),
            ("modDailyLog", "modDailyLog.bas"),
            ("modToday", "modToday.bas"),
        ]
    return shared + variant_specific


def sheet_code_behind(variant):
    """(sheet tab name, .bas file name) pairs, per variant -- injected
    directly into that sheet's own CodeModule (PLAN.md §18's sheet-named-
    .bas convention). PIC List is manager-only; Monthly Media Plan is
    employee-only; Tasks/Archive get a different file per variant
    (TasksManager.bas/ArchiveManager.bas for manager) but the SAME tab
    name either way.
    """
    is_manager = variant == "manager"
    shared = [
        ("Calendar", "Calendar.bas"),
        ("Campaigns", "Campaigns.bas"),
        ("Campaigns Archive", "Campaigns Archive.bas"),
        ("Recurring Tasks", "Recurring Tasks.bas"),
        ("Schedule Data", "Schedule Data.bas"),
        ("Schedule", "Schedule.bas"),
        ("Login Information", "Login Information.bas"),
    ]
    if is_manager:
        variant_specific = [
            ("Tasks", "TasksManager.bas"),
            ("Archive", "ArchiveManager.bas"),
            ("PIC List", "PIC List.bas"),
            ("Dashboard", "DashboardManager.bas"),
        ]
    else:
        variant_specific = [
            ("Tasks", "Tasks.bas"),
            ("Archive", "Archive.bas"),
            ("Monthly Media Plan", "Monthly Media Plan.bas"),
            ("Daily Log", "Daily Log.bas"),
            ("Dashboard", "Dashboard.bas"),
        ]
    return shared + variant_specific


def form_specs(variant):
    """UserForm specs, per variant. frmExchangeRate is Monthly-Media-
    Plan-only (confirmed via grep, plan §8) -- excluded from manager.
    """
    specs = [notes_form.SPEC]
    if variant == "employee":
        specs.append(exchange_rate_form.SPEC)
    specs += [schedule_day_form.SPEC, reminder_settings_form.SPEC, meeting_reminder_form.SPEC]
    return specs


def thisworkbook_bas_filename(variant):
    return "ThisWorkbook.bas" if variant == "employee" else "ThisWorkbookManager.bas"


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


def inject(input_xlsx, output_xlsm, variant="employee"):
    """Opens input_xlsx, injects all VBA for `variant` ("employee",
    default, unchanged from before this parameter existed, or "manager"),
    saves as a genuine macro-enabled output_xlsm (a plain .xlsx save
    would silently strip the VBA project). Raises on any failure rather
    than producing a silently partial result.
    """
    assert variant in ("employee", "manager"), f"unknown variant: {variant!r}"

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(input_xlsx)

        for comp_name, bas_name in standard_modules(variant):
            comp = wb.VBProject.VBComponents.Add(VBEXT_CT_STDMODULE)
            comp.Name = comp_name
            comp.CodeModule.AddFromString(_read_bas(bas_name))
            print(f"Injected standard module: {comp_name} (from {bas_name})")

        for spec in form_specs(variant):
            build_form_from_spec(wb, spec)
            print(f"Injected UserForm: {spec['name']}")

        for sheet_name, bas_name in sheet_code_behind(variant):
            code_name = wb.Sheets(sheet_name).CodeName
            comp = wb.VBProject.VBComponents(code_name)
            comp.CodeModule.AddFromString(_read_bas(bas_name))
            print(f"Injected {bas_name} into the {sheet_name} sheet module")

        thisworkbook_bas = thisworkbook_bas_filename(variant)
        thisworkbook_comp = _find_thisworkbook_component(wb)
        thisworkbook_comp.CodeModule.AddFromString(_read_bas(thisworkbook_bas))
        print(f"Injected {thisworkbook_bas} into the ThisWorkbook module")

        wb.SaveAs(output_xlsm, FileFormat=XL_OPENXML_WORKBOOK_MACRO_ENABLED)
        print("Saved:", output_xlsm)

        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--variant", choices=["employee", "manager"], default="employee")
    args = parser.parse_args()

    input_xlsx = output_xlsx_path(args.variant)
    output_xlsm = output_xlsm_path(args.variant)
    if not os.path.exists(input_xlsx):
        raise SystemExit(f"{input_xlsx} not found -- run build_dashboard.py --variant {args.variant} first.")
    inject(input_xlsx, output_xlsm, variant=args.variant)
