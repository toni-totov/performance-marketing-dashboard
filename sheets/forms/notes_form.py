"""frmNotes -- the shared Notes popup editor (PLAN.md §6.3), reused by
Tasks, Archive, Campaigns, Campaigns Archive, and Recurring Tasks.

A large multi-line TextBox with a pastel-yellow background, a size-toggle
button (compact <-> expanded presets), and OK/Cancel. OK writes the
TextBox content back to whichever cell the caller invoked ShowForCell(...)
with; Cancel (and Esc, via btnCancel.Cancel = True) discards edits.

Built via sheets/forms/_builder.py's build_form_from_spec() at injection
time -- see that module's docstring and PLAN.md §18 for why this isn't a
static .frm file.

Calling convention from a sheet's Worksheet_BeforeDoubleClick handler:
    frmNotes.ShowForCell target_range
ShowForCell/PrepareForCell/CommitAndClose are split into separate Public
members (rather than one do-everything Sub) specifically so the read/write
logic is testable via COM without ever triggering the blocking modal
Me.Show -- see PLAN.md §19 Phase 1 for the live test that exercises this.
"""

from ._builder import rgb_to_vba_long

_PASTEL_YELLOW = rgb_to_vba_long("FFF8C8")

SPEC = {
    "name": "frmNotes",
    "controls": [
        {
            "type": "Forms.TextBox.1",
            "name": "txtNote",
            "properties": {
                "Left": 10,
                "Top": 10,
                "Width": 250,
                "Height": 140,
                "MultiLine": True,
                "EnterKeyBehavior": True,  # MSForms default is False even with
                # MultiLine=True -- without this, Enter does nothing useful
                # instead of starting a new line (found live, PLAN.md §19 Phase 5).
                "ScrollBars": 2,  # fmScrollBarsVertical
                "BackColor": _PASTEL_YELLOW,
            },
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnResize",
            "properties": {"Caption": "⤢", "Left": 10, "Top": 160, "Width": 30, "Height": 24},
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnOK",
            "properties": {"Caption": "OK", "Left": 140, "Top": 160, "Width": 55, "Height": 24},
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnCancel",
            "properties": {
                "Caption": "Cancel",
                "Left": 200,
                "Top": 160,
                "Width": 60,
                "Height": 24,
                "Cancel": True,
            },
        },
    ],
    "code": """
Private mIsExpanded As Boolean
Private mTargetSheet As String
Private mTargetCellAddress As String

Private Sub UserForm_Initialize()
    Me.Caption = "Notes"
    ApplyCompactSize
End Sub

Private Sub ApplyCompactSize()
    Me.Width = 280
    Me.Height = 220
    txtNote.Width = 250
    txtNote.Height = 140
    PositionButtons
    mIsExpanded = False
End Sub

Private Sub ApplyExpandedSize()
    Me.Width = 480
    Me.Height = 420
    txtNote.Width = 450
    txtNote.Height = 330
    PositionButtons
    mIsExpanded = True
End Sub

Private Sub PositionButtons()
    Dim footerTop As Single
    footerTop = txtNote.Top + txtNote.Height + 10
    btnResize.Top = footerTop
    btnOK.Top = footerTop
    btnCancel.Top = footerTop
    btnOK.Left = txtNote.Left + txtNote.Width - 120
    btnCancel.Left = txtNote.Left + txtNote.Width - 60
End Sub

Private Sub btnResize_Click()
    If mIsExpanded Then
        ApplyCompactSize
    Else
        ApplyExpandedSize
    End If
End Sub

Public Sub ShowForCell(targetCell As Range)
    PrepareForCell targetCell
    Me.Show
End Sub

Public Sub PrepareForCell(targetCell As Range)
    mTargetSheet = targetCell.Worksheet.Name
    mTargetCellAddress = targetCell.Address
    txtNote.Text = CStr(targetCell.Value)
End Sub

Private Sub btnOK_Click()
    CommitAndClose
End Sub

Public Sub CommitAndClose()
    ThisWorkbook.Worksheets(mTargetSheet).Range(mTargetCellAddress).Value = txtNote.Text
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
""",
}
