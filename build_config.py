"""Shared variant configuration for build_dashboard.py and inject_vba.py
(PLAN.md, "Manager-variant workbook" §1). A single source of truth for
"what does each variant produce" -- the two build scripts import from here
rather than each hardcoding output filenames.
"""

import os

_REPO_ROOT = os.path.dirname(os.path.abspath(__file__))

VARIANTS = {
    "employee": {"output_basename": "Performance Marketing Dashboard"},
    "manager": {"output_basename": "Task Tracker - Manager"},
}


def output_xlsx_path(variant):
    return os.path.join(_REPO_ROOT, f"{VARIANTS[variant]['output_basename']}.xlsx")


def output_xlsm_path(variant):
    return os.path.join(_REPO_ROOT, f"{VARIANTS[variant]['output_basename']}.xlsm")
