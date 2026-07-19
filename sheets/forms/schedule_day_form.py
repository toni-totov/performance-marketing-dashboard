"""frmScheduleDay -- the Schedule day popup (PLAN.md §13.3). Opened by
double-clicking a date cell on Schedule. Three sections: a Work Location
dropdown (upserts that date's row in Schedule Data), a read-list of that
date's meetings (one-time exact-date matches + Active recurring matches)
with a single shared "Remove Selected" button (confirmed with the user:
a ListBox + one shared Remove button, not a dynamically-generated button
per row -- simpler and more robust for a low-stakes personal-scheduling
feature), and an "Add Meeting" mini-form.

Built via sheets/forms/_builder.py's build_form_from_spec() at injection
time -- see that module's docstring and PLAN.md §18 for why this isn't a
static .frm file.

Calling convention from Schedule's Worksheet_BeforeDoubleClick handler:
    frmScheduleDay.ShowForDate clickedDate
ShowForDate/PrepareForDate are split the same way as frmNotes' ShowForCell/
PrepareForCell (Phase 1). The three actions (save location, remove a
meeting, add a meeting) are each split into a MsgBox-free Public core
(SaveWorkLocation always succeeds; RemoveSelectedMeeting has nothing to
validate; TryAddMeeting returns False without writing anything if the
Meeting Name is blank) separate from their thin _Click wrappers, so the
real logic is testable via Application.Run without ever triggering
Me.Show or a validation MsgBox -- same reasoning as Tasks' CascadeDelete/
WouldCreateCycle split (PLAN.md §19 Phase 3). The actual Schedule Data
table reads/writes (AddMeeting/DeleteMeetingByKey/NextMeetingKey/
MeetingOccursOn) live in modSchedule.bas, not here, since
ScheduleReminders() (§13.5) needs the same "does meeting X occur on date
D" check and must not duplicate it.
"""

SPEC = {
    "name": "frmScheduleDay",
    "controls": [
        {
            "type": "Forms.Label.1",
            "name": "lblDate",
            "properties": {"Caption": "", "Left": 10, "Top": 10, "Width": 260, "Height": 18},
        },
        {
            "type": "Forms.Label.1",
            "name": "lblLocation",
            "properties": {"Caption": "Work Location:", "Left": 10, "Top": 38, "Width": 90, "Height": 18},
        },
        {
            "type": "Forms.ComboBox.1",
            "name": "cboWorkLocation",
            "properties": {"Left": 100, "Top": 36, "Width": 110, "Height": 20, "Style": 2},
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnSaveLocation",
            "properties": {"Caption": "Save", "Left": 215, "Top": 35, "Width": 55, "Height": 22},
        },
        {
            "type": "Forms.Label.1",
            "name": "lblMeetingsHeader",
            "properties": {"Caption": "Meetings on this day:", "Left": 10, "Top": 68, "Width": 200, "Height": 18},
        },
        {
            "type": "Forms.ListBox.1",
            "name": "lstMeetings",
            "properties": {"Left": 10, "Top": 88, "Width": 260, "Height": 80},
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnRemoveSelected",
            "properties": {"Caption": "Remove Selected", "Left": 10, "Top": 172, "Width": 130, "Height": 22},
        },
        {
            "type": "Forms.Label.1",
            "name": "lblAddHeader",
            "properties": {"Caption": "Add Meeting:", "Left": 10, "Top": 204, "Width": 150, "Height": 18},
        },
        {
            "type": "Forms.Label.1",
            "name": "lblName",
            "properties": {"Caption": "Name:", "Left": 10, "Top": 226, "Width": 50, "Height": 18},
        },
        {
            "type": "Forms.TextBox.1",
            "name": "txtMeetingName",
            "properties": {"Left": 65, "Top": 224, "Width": 205, "Height": 20},
        },
        {
            "type": "Forms.Label.1",
            "name": "lblStart",
            "properties": {"Caption": "Start:", "Left": 10, "Top": 250, "Width": 50, "Height": 18},
        },
        {
            "type": "Forms.ComboBox.1",
            "name": "cboStartHour",
            "properties": {"Left": 65, "Top": 248, "Width": 50, "Height": 20, "Style": 2},
        },
        {
            "type": "Forms.ComboBox.1",
            "name": "cboStartMinute",
            "properties": {"Left": 120, "Top": 248, "Width": 50, "Height": 20, "Style": 2},
        },
        {
            "type": "Forms.Label.1",
            "name": "lblEnd",
            "properties": {"Caption": "End:", "Left": 10, "Top": 274, "Width": 50, "Height": 18},
        },
        {
            "type": "Forms.ComboBox.1",
            "name": "cboEndHour",
            "properties": {"Left": 65, "Top": 272, "Width": 50, "Height": 20, "Style": 2},
        },
        {
            "type": "Forms.ComboBox.1",
            "name": "cboEndMinute",
            "properties": {"Left": 120, "Top": 272, "Width": 50, "Height": 20, "Style": 2},
        },
        {
            "type": "Forms.Label.1",
            "name": "lblRecurrence",
            "properties": {"Caption": "Recurrence:", "Left": 10, "Top": 298, "Width": 70, "Height": 18},
        },
        {
            "type": "Forms.ComboBox.1",
            "name": "cboRecurrence",
            "properties": {"Left": 90, "Top": 296, "Width": 110, "Height": 20, "Style": 2},
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnAddMeeting",
            "properties": {"Caption": "Add Meeting", "Left": 10, "Top": 324, "Width": 110, "Height": 22},
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnClose",
            "properties": {"Caption": "Close", "Left": 215, "Top": 354, "Width": 55, "Height": 24, "Cancel": True},
        },
    ],
    "code": """
Private mTargetDate As Date

Private Sub UserForm_Initialize()
    Me.Caption = "Schedule Day"
    Me.Width = 300
    Me.Height = 415

    lstMeetings.ColumnCount = 2
    lstMeetings.ColumnWidths = "220 pt;0 pt"

    Dim h As Long
    For h = 0 To 23
        cboStartHour.AddItem Format(h, "00")
        cboEndHour.AddItem Format(h, "00")
    Next h
    Dim mnt As Variant
    For Each mnt In Array("00", "15", "30", "45")
        cboStartMinute.AddItem mnt
        cboEndMinute.AddItem mnt
    Next mnt
    cboStartHour.Value = "09"
    cboStartMinute.Value = "00"
    cboEndHour.Value = "10"
    cboEndMinute.Value = "00"

    cboWorkLocation.AddItem ""
    cboWorkLocation.AddItem "Home Office"
    cboWorkLocation.AddItem "Office"
    cboWorkLocation.AddItem "Holiday"
    ' "Sick"/"Out of Office" (user-requested, PLAN.md §19 Phase 5): two
    ' more away-states alongside Holiday -- all three exclude that date
    ' from recurring meeting occurrences (modSchedule.bas's
    ' MeetingOccursOn / this sheet's own badge formula).
    cboWorkLocation.AddItem "Sick"
    cboWorkLocation.AddItem "Out of Office"

    cboRecurrence.AddItem "One-time"
    cboRecurrence.AddItem "Daily"
    ' "Weekdays" (user-requested, PLAN.md §19 Phase 5): a separate
    ' recurrence type from "Daily" (Monday-Friday only), not a modifier
    ' on it -- existing "Daily" meetings keep their exact current
    ' every-day-including-weekends meaning unchanged.
    cboRecurrence.AddItem "Weekdays"
    cboRecurrence.AddItem "Weekly"
    cboRecurrence.AddItem "Monthly"
    cboRecurrence.Value = "One-time"
End Sub

Public Sub ShowForDate(targetDate As Date)
    PrepareForDate targetDate
    Me.Show
End Sub

Public Sub PrepareForDate(targetDate As Date)
    mTargetDate = targetDate
    ' Locale-proof weekday/month names (PLAN.md §19 Phase 5, found live) --
    ' see modDesignSystem.bas's EnglishMonthName/EnglishWeekdayName for
    ' why Format(...,"dddd mmmm") isn't safe (renders in system locale).
    lblDate.Caption = EnglishWeekdayName(targetDate) & ", " & Format(targetDate, "dd") & _
        " " & EnglishMonthName(Month(targetDate)) & " " & Format(targetDate, "yyyy")
    txtMeetingName.Text = ""
    LoadWorkLocation
    LoadMeetings
End Sub

Private Sub LoadWorkLocation()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Schedule Data")
    Dim lastRow As Long
    lastRow = modHierarchy.LastHierarchyRow(ws, 1, 2)

    cboWorkLocation.Value = ""
    Dim r As Long
    For r = 2 To lastRow
        If Len(Trim(ws.Cells(r, 1).Value & "")) > 0 Then
            If CLng(ws.Cells(r, 1).Value) = CLng(mTargetDate) Then
                cboWorkLocation.Value = ws.Cells(r, 2).Value
                Exit For
            End If
        End If
    Next r
End Sub

Public Sub SaveWorkLocation()
    modSchedule.UpsertWorkLocation mTargetDate, cboWorkLocation.Value
End Sub

Private Sub LoadMeetings()
    lstMeetings.Clear

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Schedule Data")
    Dim lastRow As Long
    lastRow = modHierarchy.LastHierarchyRow(ws, 10, 2)

    Dim r As Long
    For r = 2 To lastRow
        If Len(Trim(ws.Cells(r, 10).Value & "")) > 0 Then
            If modSchedule.MeetingOccursOn(ws, r, mTargetDate) Then
                Dim displayText As String
                displayText = Format(ws.Cells(r, 5).Value, "hh:nn") & "-" & Format(ws.Cells(r, 6).Value, "hh:nn") & "  " & ws.Cells(r, 4).Value
                lstMeetings.AddItem displayText
                lstMeetings.List(lstMeetings.ListCount - 1, 1) = CStr(ws.Cells(r, 10).Value)
            End If
        End If
    Next r
End Sub

Public Sub RemoveSelectedMeeting()
    If lstMeetings.ListIndex < 0 Then Exit Sub
    Dim meetingKey As String
    meetingKey = lstMeetings.List(lstMeetings.ListIndex, 1)
    modSchedule.DeleteMeetingByKey CDbl(meetingKey)
    LoadMeetings
    modSchedule.ScheduleReminders
End Sub

' No MsgBox here -- returns False (nothing written) if the meeting name
' is blank; see module docstring for why this is split from btnAddMeeting_Click.
Public Function TryAddMeeting() As Boolean
    Dim meetingName As String
    meetingName = Trim(txtMeetingName.Text)
    If Len(meetingName) = 0 Then
        TryAddMeeting = False
        Exit Function
    End If

    Dim startTime As Date, endTime As Date
    startTime = TimeSerial(CLng(cboStartHour.Value), CLng(cboStartMinute.Value), 0)
    endTime = TimeSerial(CLng(cboEndHour.Value), CLng(cboEndMinute.Value), 0)

    modSchedule.AddMeeting meetingName, startTime, endTime, cboRecurrence.Value, mTargetDate

    txtMeetingName.Text = ""
    LoadMeetings
    modSchedule.ScheduleReminders

    TryAddMeeting = True
End Function

Private Sub btnSaveLocation_Click()
    SaveWorkLocation
End Sub

Private Sub btnRemoveSelected_Click()
    RemoveSelectedMeeting
End Sub

Private Sub btnAddMeeting_Click()
    If Not TryAddMeeting() Then
        MsgBox "Please enter a meeting name.", vbExclamation, "Meeting Name Required"
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
""",
}
