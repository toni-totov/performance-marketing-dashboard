"""frmReminderSettings -- the Schedule reminder-settings popup (PLAN.md
§13.5). Three numeric TextBoxes (First Offset minutes / Count / Interval
minutes), global (not per-meeting, matching "I would be able to set,"
singular). OK validates and writes ReminderFirstOffset/ReminderCount/
ReminderInterval, then reschedules reminders immediately so the change
takes effect without waiting for the next Workbook_Open; Cancel discards.

Built via sheets/forms/_builder.py's build_form_from_spec() at injection
time -- see that module's docstring and PLAN.md §18.

Calling convention from Schedule's Worksheet_BeforeDoubleClick handler:
    frmReminderSettings.ShowSettings
Same Public-for-modal-bypass split as frmExchangeRate: TryApplySettings()
is the MsgBox-free, testable core (valid input writes and returns True;
invalid input leaves the named cells untouched and returns False);
btnOK_Click is the thin MsgBox-showing production wrapper.
"""

SPEC = {
    "name": "frmReminderSettings",
    "controls": [
        {
            "type": "Forms.Label.1",
            "name": "lblFirstOffset",
            "properties": {"Caption": "First reminder (min before):", "Left": 10, "Top": 12, "Width": 160, "Height": 18},
        },
        {
            "type": "Forms.TextBox.1",
            "name": "txtFirstOffset",
            "properties": {"Left": 175, "Top": 10, "Width": 50, "Height": 20},
        },
        {
            "type": "Forms.Label.1",
            "name": "lblCount",
            "properties": {"Caption": "Number of reminders:", "Left": 10, "Top": 38, "Width": 160, "Height": 18},
        },
        {
            "type": "Forms.TextBox.1",
            "name": "txtCount",
            "properties": {"Left": 175, "Top": 36, "Width": 50, "Height": 20},
        },
        {
            "type": "Forms.Label.1",
            "name": "lblInterval",
            "properties": {"Caption": "Interval between (min):", "Left": 10, "Top": 64, "Width": 160, "Height": 18},
        },
        {
            "type": "Forms.TextBox.1",
            "name": "txtInterval",
            "properties": {"Left": 175, "Top": 62, "Width": 50, "Height": 20},
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnOK",
            "properties": {"Caption": "OK", "Left": 90, "Top": 95, "Width": 55, "Height": 24},
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnCancel",
            "properties": {"Caption": "Cancel", "Left": 150, "Top": 95, "Width": 60, "Height": 24, "Cancel": True},
        },
    ],
    "code": """
Private Sub UserForm_Initialize()
    Me.Caption = "Reminder Settings"
    Me.Width = 250
    Me.Height = 165
End Sub

Public Sub ShowSettings()
    PrepareSettings
    Me.Show
End Sub

Public Sub PrepareSettings()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Schedule")
    txtFirstOffset.Text = CStr(ws.Range("ReminderFirstOffset").Value)
    txtCount.Text = CStr(ws.Range("ReminderCount").Value)
    txtInterval.Text = CStr(ws.Range("ReminderInterval").Value)
End Sub

' No MsgBox here -- returns False (named cells untouched) on invalid
' input; see module docstring for why this is split from btnOK_Click.
Public Function TryApplySettings() As Boolean
    If Not IsNumeric(txtFirstOffset.Text) Or Not IsNumeric(txtCount.Text) Or Not IsNumeric(txtInterval.Text) Then
        TryApplySettings = False
        Exit Function
    End If

    Dim fo As Double, cnt As Double, iv As Double
    fo = CDbl(txtFirstOffset.Text)
    cnt = CDbl(txtCount.Text)
    iv = CDbl(txtInterval.Text)

    If fo < 0 Or cnt < 1 Or iv < 0 Then
        TryApplySettings = False
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Schedule")
    ws.Range("ReminderFirstOffset").Value = fo
    ws.Range("ReminderCount").Value = cnt
    ws.Range("ReminderInterval").Value = iv

    modSchedule.ScheduleReminders

    TryApplySettings = True
End Function

Private Sub btnOK_Click()
    If TryApplySettings() Then
        Unload Me
    Else
        MsgBox "Please enter valid settings (First Offset >= 0, Count >= 1, Interval >= 0).", vbExclamation, "Invalid Settings"
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
""",
}
