"""frmMeetingReminder -- the popup shown when a scheduled meeting
reminder fires (PLAN.md §13.5). Not a bare MsgBox, for consistency with
the rest of the workbook's polish. Shows meeting name, start time, and
minutes-until-start; OK dismisses (no Cancel -- there's nothing to
discard, it's purely informational).

Built via sheets/forms/_builder.py's build_form_from_spec() at injection
time -- see that module's docstring and PLAN.md §18.

Calling convention (from modSchedule.bas's FireMeetingReminder, the
Application.OnTime callback target):
    frmMeetingReminder.ShowReminder meetingName, startTimeText, minutesUntil
ShowReminder's own Me.Show is never exercised directly in automated
testing (Application.OnTime callbacks and modal popups can't be driven
headlessly) -- modSchedule.bas's PopNextReminderInfo() is what's tested
instead, verifying the right reminder would be shown without ever
actually showing it.
"""

SPEC = {
    "name": "frmMeetingReminder",
    "controls": [
        {
            "type": "Forms.Label.1",
            "name": "lblMessage",
            "properties": {"Caption": "", "Left": 10, "Top": 10, "Width": 220, "Height": 60},
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnOK",
            "properties": {"Caption": "OK", "Left": 90, "Top": 75, "Width": 60, "Height": 24, "Default": True},
        },
    ],
    "code": """
Private Sub UserForm_Initialize()
    Me.Caption = "Meeting Reminder"
    Me.Width = 250
    Me.Height = 140
End Sub

Public Sub ShowReminder(meetingName As String, startTimeText As String, minutesUntil As Long)
    PrepareReminder meetingName, startTimeText, minutesUntil
    Me.Show
End Sub

Public Sub PrepareReminder(meetingName As String, startTimeText As String, minutesUntil As Long)
    Dim whenText As String
    If minutesUntil = 0 Then
        whenText = "starting now"
    Else
        whenText = "in " & minutesUntil & " minute(s)"
    End If
    lblMessage.Caption = meetingName & vbCrLf & "Starts at " & startTimeText & " (" & whenText & ")"
End Sub

Private Sub btnOK_Click()
    Unload Me
End Sub
""",
}
