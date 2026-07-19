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

import os

import openpyxl

from sheets import (
    archive,
    calendar,
    campaigns,
    campaigns_archive,
    dashboard,
    instructions,
    login_information,
    monthly_media_plan,
    quick_links,
    recurring_tasks,
    schedule,
    schedule_data,
    tasks,
    time_log,
)

OUTPUT_XLSX = os.path.join(os.path.dirname(__file__), "Performance Marketing Dashboard.xlsx")


def build():
    """Builds the full workbook structure (no VBA -- see inject_vba.py)
    and saves it to OUTPUT_XLSX. Returns the openpyxl Workbook.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    time_log.build(wb)
    tasks.build(wb)
    archive.build(wb)
    campaigns.build(wb)
    campaigns_archive.build(wb)
    recurring_tasks.build(wb)
    calendar.build(wb)
    monthly_media_plan.build(wb)
    schedule_data.build(wb)
    schedule.build(wb)
    login_information.build(wb)
    quick_links.build(wb)
    instructions.build(wb)
    dashboard.build(wb)

    wb.save(OUTPUT_XLSX)
    return wb


if __name__ == "__main__":
    build()
    print("Saved:", OUTPUT_XLSX)
