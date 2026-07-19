"""Generic UserForm builder: constructs a real VBA UserForm via VBIDE COM
automation from a plain-data spec.

PLAN.md §18's correction: hand-authored .frm text without a real,
Designer-backed .frx does not produce a working form -- confirmed via
live testing, it silently misclassifies the component and breaks
Application.Run for the whole VBA project. Building via
VBComponents.Add(3) + .Designer.Controls.Add + .CodeModule.AddFromString
does work and was confirmed to survive save/close/reopen.

Each form's own module (e.g. notes_form.py) defines a SPEC dict; this
module turns any such spec into a real form component in an open
Workbook's VBA project. Used by inject_vba.py (Phase 4) for every
UserForm in the workbook -- written once here, not duplicated per form.

Spec shape:
{
    "name": "frmNotes",
    "controls": [
        {"type": "Forms.TextBox.1", "name": "txtNote", "properties": {...}},
        ...
    ],
    "code": "<VBA code-behind text>",
}

Note: a UserForm's own Width/Height/Caption are NOT set via COM at build
time -- Designer.Properties(...) is not reachable through this automation
path (confirmed via testing: fails with "<unknown>.Properties" on the
gencache-wrapped Designer object). They're set instead by the form's own
VBA code (Me.Width = ... in UserForm_Initialize), which is plain text
baked into "code" above. Control-level properties (Left, Top, Width,
Height, Caption, MultiLine, BackColor, Cancel, etc.) ARE settable directly
via COM attribute assignment on the control object and are applied here.
"""

# vbext_ComponentType.vbext_ct_MSForm
_VBEXT_CT_MSFORM = 3


def build_form_from_spec(wb, spec):
    """Build one UserForm from `spec` in the given open Workbook's VBA
    project (wb.VBProject). Returns the created VBComponent.
    """
    form_comp = wb.VBProject.VBComponents.Add(_VBEXT_CT_MSFORM)
    form_comp.Name = spec["name"]

    designer = form_comp.Designer
    for ctrl_spec in spec.get("controls", []):
        ctrl = designer.Controls.Add(ctrl_spec["type"], ctrl_spec["name"], True)
        for prop_name, prop_value in ctrl_spec.get("properties", {}).items():
            setattr(ctrl, prop_name, prop_value)

    code = spec.get("code", "").strip()
    if code:
        form_comp.CodeModule.AddFromString(code)

    return form_comp


def rgb_to_vba_long(hex_str):
    """Convert a "#RRGGBB" or "RRGGBB" hex string to the BGR-ordered Long
    VBA color values require (same conversion used to generate
    vba/modDesignSystem.bas's RGB() constants -- see sheets/design_system.py).
    """
    hex_str = hex_str.lstrip("#")
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return r + g * 256 + b * 65536
