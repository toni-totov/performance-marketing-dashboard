"""Orchestrator entry point — builds the workbook's structure via openpyxl.

See PLAN.md for the full spec, and PLAN.md §19 (Implementation Checklist)
for build order. VBA is injected separately by inject_vba.py.

Build order follows §2's dependency chain (Tasks -> Archive -> Campaigns
-> Campaigns Archive -> Recurring Tasks -> Schedule Data -> Schedule),
with each explicit cross-sheet dependency confirmed by the corresponding
sheet module's own docstring: campaigns.build() writes directly onto
Tasks' own Linked Campaign dropdown (completing the stub Tasks' own build
leaves behind) so Tasks must exist first; recurring_tasks.build() sources
its own Linked Campaign dropdown from Campaigns' CampaignNameList; and
schedule.build() reads Schedule Data's raw storage layout directly.
Calendar and Monthly Media Plan have no hard build-order dependency of
their own. Dashboard's formulas reference nearly every other sheet, so it
is built last.
"""

import argparse

import openpyxl

from build_config import output_xlsx_path
from sheets import (
    archive,
    calendar,
    campaigns,
    campaigns_archive,
    daily_log,
    dashboard,
    instructions,
    login_information,
    monthly_media_plan,
    pic_list,
    quick_links,
    recurring_tasks,
    schedule,
    schedule_data,
    tasks,
    time_log,
)
from sheets import conditional_formatting as cf

# Kept for backward compatibility with any external reference to the old
# unconditional (employee-only) module-level constant.
OUTPUT_XLSX = output_xlsx_path("employee")


def build(variant="employee"):
    """Builds the full workbook structure (no VBA -- see inject_vba.py)
    and saves it to build_config.output_xlsx_path(variant). Returns the
    openpyxl Workbook. `variant="employee"` (default) preserves the exact
    behavior this function had before the parameter existed -- confirmed
    byte-identical output.

    Manager (PLAN.md, "Manager-variant workbook" §1): skips Time Log and
    Monthly Media Plan entirely (neither exists for that variant), and
    builds the PIC List roster sheet (manager-only, feeds Tasks' PIC
    dropdown) before Tasks itself so the dropdown's source range already
    exists when tasks.build() wires it up.
    """
    assert variant in ("employee", "manager"), f"unknown variant: {variant!r}"
    is_manager = variant == "manager"

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    if not is_manager:
        time_log.build(wb)
    if is_manager:
        pic_list.build(wb)
    tasks.build(wb, variant=variant)
    archive.build(wb, variant=variant)
    if not is_manager:
        daily_log.build(wb)  # V4, PLAN.md §22.6 -- a view over Archive, placed right after it
    campaigns.build(wb)
    campaigns_archive.build(wb)
    recurring_tasks.build(wb)
    calendar.build(wb)
    if not is_manager:
        monthly_media_plan.build(wb)
    schedule_data.build(wb)
    schedule.build(wb)
    login_information.build(wb)
    quick_links.build(wb)
    instructions.build(wb)
    dashboard.build(wb, variant=variant)

    # Last, after every sheet has added its rules (PLAN.md §22.2 #2).
    cf.demote_banding_rules(wb)

    output_path = output_xlsx_path(variant)
    wb.save(output_path)
    return wb


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--variant", choices=["employee", "manager"], default="employee")
    args = parser.parse_args()

    build(variant=args.variant)
    print("Saved:", output_xlsx_path(args.variant))
