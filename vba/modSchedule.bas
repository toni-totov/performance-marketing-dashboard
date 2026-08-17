Attribute VB_Name = "modSchedule"
Option Explicit

' Schedule + Schedule Data business logic (PLAN.md §13). Three groups:
'   1. Month navigation for Schedule's own ScheduleYear/ScheduleMonth --
'      mirrors modCalendar.bas's shape but on Schedule's own named ranges
'      (can't share instances, same reasoning as Recurring Tasks not
'      sharing modTasks.bas's hierarchy functions).
'   2. Schedule Data read/write helpers (UpsertWorkLocation/AddMeeting/
'      DeleteMeetingByKey/MeetingOccursOn) -- shared by frmScheduleDay
'      (the popup that edits Schedule Data) AND the reminder engine below
'      (which needs the same "does meeting X occur on date D" check), so
'      they live here rather than being duplicated in the form's own
'      code-behind.
'   3. The reminder engine (§13.5) -- CollectDueReminders is a pure,
'      MsgBox-free date-math function (testable directly); ScheduleReminders
'      is the production entry point that cancels any previously scheduled
'      Application.OnTime calls (tracked in the module-level
'      mScheduledTimes/mReminderQueue Collections -- an in-memory list,
'      not a sheet-based one, since a scheduled OnTime call itself can't
'      survive a file close/reopen either, so persisting the tracking
'      list to a sheet would track state that's already gone the moment
'      it would matter) and reschedules fresh ones. FireMeetingReminder
'      is the fixed callback name every Application.OnTime call points
'      at (OnTime cannot pass parameters to its target Sub); it pops the
'      next entry off a FIFO queue built in the same chronological order
'      the reminders were scheduled in, so "the next OnTime callback to
'      actually fire" and "the next queue entry" always correspond --
'      this only holds because CollectDueReminders sorts all reminders
'      (across every meeting) by time before either the queue or the
'      Application.OnTime calls are populated; scheduling them in the
'      Meetings table's row order instead would desync the two.

' --- Schedule Data column layout (sheets/schedule_data.py) -----------------
Private Const WL_COL_DATE As Long = 1
Private Const WL_COL_LOCATION As Long = 2
Private Const WL_FIRST_DATA_ROW As Long = 2

Private Const MT_COL_MEETING_NAME As Long = 4
Private Const MT_COL_START_TIME As Long = 5
Private Const MT_COL_END_TIME As Long = 6
Private Const MT_COL_RECURRENCE As Long = 7
Private Const MT_COL_ANCHOR_DATE As Long = 8
Private Const MT_COL_STATUS As Long = 9
Private Const MT_COL_KEY As Long = 10
Private Const MT_FIRST_DATA_ROW As Long = 2

' In-memory tracking for the reminder engine -- see module header comment.
Private mScheduledTimes As Collection
Private mReminderQueue As Collection

' Field separator for the FIFO queue's packed "Name<sep>Time<sep>Minutes"
' strings. The null character (a non-typable control character), not a
' literal "|" -- a "|" can legitimately appear in a user-typed meeting
' name, which used to shift every field after it and throw inside
' FireMeetingReminder (user-reported: an unhandled error firing as a
' blocking modal the moment a reminder popped).
Private Const REMINDER_INFO_SEP As String = vbNullChar


' ============================================================================
' Month navigation (Schedule's own ScheduleYear/ScheduleMonth)
' ============================================================================

Private Function ScheduleSheet() As Worksheet
    Set ScheduleSheet = ThisWorkbook.Sheets("Schedule")
End Function

Public Sub GoToPrevMonth()
    Dim ws As Worksheet
    Set ws = ScheduleSheet()
    Dim y As Long, m As Long
    y = ws.Range("ScheduleYear").Value
    m = ws.Range("ScheduleMonth").Value
    m = m - 1
    If m < 1 Then
        m = 12
        y = y - 1
    End If
    ws.Range("ScheduleYear").Value = y
    ws.Range("ScheduleMonth").Value = m
End Sub

Public Sub GoToNextMonth()
    Dim ws As Worksheet
    Set ws = ScheduleSheet()
    Dim y As Long, m As Long
    y = ws.Range("ScheduleYear").Value
    m = ws.Range("ScheduleMonth").Value
    m = m + 1
    If m > 12 Then
        m = 1
        y = y + 1
    End If
    ws.Range("ScheduleYear").Value = y
    ws.Range("ScheduleMonth").Value = m
End Sub

Public Sub GoToToday()
    Dim ws As Worksheet
    Set ws = ScheduleSheet()
    ws.Range("ScheduleYear").Value = Year(Date)
    ws.Range("ScheduleMonth").Value = Month(Date)
End Sub

' ============================================================================
' Schedule Data read/write helpers
' ============================================================================

' Upserts targetDate's Work Location row -- a blank location clears (removes)
' that date's row entirely rather than leaving a dangling Date-with-no-
' Location entry.
Public Sub UpsertWorkLocation(targetDate As Date, locationText As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Schedule Data")

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim lastRow As Long
    lastRow = modHierarchy.LastHierarchyRow(ws, WL_COL_DATE, WL_FIRST_DATA_ROW)
    Dim foundRow As Long
    foundRow = 0
    Dim r As Long
    For r = WL_FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, WL_COL_DATE).Value & "")) > 0 Then
            If CLng(ws.Cells(r, WL_COL_DATE).Value) = CLng(targetDate) Then
                foundRow = r
                Exit For
            End If
        End If
    Next r

    If foundRow > 0 Then
        If Len(Trim$(locationText)) = 0 Then
            ws.Cells(foundRow, WL_COL_DATE).Value = ""
            ws.Cells(foundRow, WL_COL_LOCATION).Value = ""
        Else
            ws.Cells(foundRow, WL_COL_LOCATION).Value = locationText
        End If
    ElseIf Len(Trim$(locationText)) > 0 Then
        Dim tbl As ListObject
        Set tbl = ws.ListObjects("WorkLocationTable")
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add
        newRow.Range(1, 1).Value = targetDate
        newRow.Range(1, 2).Value = locationText
    End If

    Application.EnableEvents = prevEvents
End Sub

Public Sub AddMeeting(meetingName As String, startTime As Date, endTime As Date, recurrence As String, anchorDate As Date)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Schedule Data")

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim tbl As ListObject
    Set tbl = ws.ListObjects("MeetingsTable")
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add

    newRow.Range(1, 1).Value = meetingName
    newRow.Range(1, 2).Value = startTime
    newRow.Range(1, 3).Value = endTime
    newRow.Range(1, 4).Value = recurrence
    newRow.Range(1, 5).Value = anchorDate
    newRow.Range(1, 6).Value = "Active"
    newRow.Range(1, 7).Value = NextMeetingKey()

    Application.EnableEvents = prevEvents
End Sub

Private Function NextMeetingKey() As Double
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Schedule Data")
    Dim maxKey As Double
    maxKey = 0
    Dim lastRow As Long, r As Long
    lastRow = modHierarchy.LastHierarchyRow(ws, MT_COL_KEY, MT_FIRST_DATA_ROW)
    For r = MT_FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, MT_COL_KEY).Value & "")) > 0 Then
            If CDbl(ws.Cells(r, MT_COL_KEY).Value) > maxKey Then maxKey = CDbl(ws.Cells(r, MT_COL_KEY).Value)
        End If
    Next r
    NextMeetingKey = maxKey + 1
End Function

' No per-occurrence state exists for a recurring meeting to remove just
' one instance of (§13.4) -- Remove always deletes the whole rule row,
' whether One-time or recurring. Deletes via the MeetingsTable ListObject's
' own ListRows, NOT a raw Rows(r).Delete -- Schedule Data has two
' independent Tables (WorkLocationTable in A:B, MeetingsTable in D:J)
' sharing the same physical row-number space in different columns, so a
' whole-row delete would also shift/corrupt whatever WorkLocationTable
' data happens to sit on that same row. Confirmed via COM: a raw
' Rows(r).Delete errors outright ("Delete method of Range class failed")
' rather than silently doing the wrong thing, since Excel refuses a
' full-row delete that would partially restructure an unrelated table.
Public Sub DeleteMeetingByKey(meetingKey As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Schedule Data")

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim tbl As ListObject
    Set tbl = ws.ListObjects("MeetingsTable")

    Dim lastRow As Long
    lastRow = modHierarchy.LastHierarchyRow(ws, MT_COL_KEY, MT_FIRST_DATA_ROW)
    Dim r As Long
    For r = MT_FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, MT_COL_KEY).Value & "")) > 0 Then
            If CDbl(ws.Cells(r, MT_COL_KEY).Value) = meetingKey Then
                tbl.ListRows(r - MT_FIRST_DATA_ROW + 1).Delete
                Exit For
            End If
        End If
    Next r

    Application.EnableEvents = prevEvents
End Sub

' True if targetDate's Work Location is Holiday/Sick/Out of Office (an
' "away" state, user-requested, PLAN.md §19 Phase 5) -- used by
' MeetingOccursOn to exclude recurring meetings from a day the user
' won't actually be working. Same linear-scan lookup pattern as
' UpsertWorkLocation's own.
Private Function IsExcludedDay(targetDate As Date) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Schedule Data")

    Dim lastRow As Long
    lastRow = modHierarchy.LastHierarchyRow(ws, WL_COL_DATE, WL_FIRST_DATA_ROW)

    Dim r As Long
    For r = WL_FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, WL_COL_DATE).Value & "")) > 0 Then
            If CLng(ws.Cells(r, WL_COL_DATE).Value) = CLng(targetDate) Then
                Dim loc As String
                loc = Trim$(ws.Cells(r, WL_COL_LOCATION).Value & "")
                IsExcludedDay = (loc = "Holiday" Or loc = "Sick" Or loc = "Out of Office")
                Exit Function
            End If
        End If
    Next r

    IsExcludedDay = False
End Function

' Does the meeting rule on row r occur on targetDate? One-time: exact
' Anchor Date match (regardless of Status -- a one-time meeting has no
' Active/Inactive concept of its own occurrence, matching §13.1's table).
' Recurring: only while Active, weekday/day-of-month pattern from Anchor
' Date, AND targetDate isn't an excluded (Holiday/Sick/Out of Office) day
' -- One-time meetings are unaffected by exclusion, since an explicitly-
' picked exact date is a deliberate choice, not a projected pattern.
' Shared by frmScheduleDay's LoadMeetings and CollectDueReminders below
' -- see module header comment for why this isn't duplicated.
Public Function MeetingOccursOn(ws As Worksheet, r As Long, targetDate As Date) As Boolean
    Dim recurrence As String, status As String
    recurrence = Trim$(ws.Cells(r, MT_COL_RECURRENCE).Value & "")
    status = Trim$(ws.Cells(r, MT_COL_STATUS).Value & "")

    If recurrence = "One-time" Then
        MeetingOccursOn = (CLng(ws.Cells(r, MT_COL_ANCHOR_DATE).Value) = CLng(targetDate))
    ElseIf status = "Active" Then
        Dim anchorDate As Date
        anchorDate = ws.Cells(r, MT_COL_ANCHOR_DATE).Value
        Select Case recurrence
            Case "Daily"
                MeetingOccursOn = True
            Case "Weekdays"
                ' User-requested addition (PLAN.md §19 Phase 5): a
                ' separate recurrence type from "Daily" (which keeps its
                ' exact original every-day-including-weekends meaning),
                ' restricted to Monday-Friday. vbMonday-based Weekday()
                ' returns 1=Monday..7=Sunday, so <=5 covers Mon-Fri.
                MeetingOccursOn = (Weekday(targetDate, vbMonday) <= 5)
            Case "Weekly"
                MeetingOccursOn = (Weekday(anchorDate, vbMonday) = Weekday(targetDate, vbMonday))
            Case "Monthly"
                MeetingOccursOn = (Day(anchorDate) = Day(targetDate))
            Case Else
                MeetingOccursOn = False
        End Select

        If MeetingOccursOn And IsExcludedDay(targetDate) Then
            MeetingOccursOn = False
        End If
    Else
        MeetingOccursOn = False
    End If
End Function

' ============================================================================
' Reminders (§13.5)
' ============================================================================

' Computes every due reminder for todayDate's meetings, sorted ascending
' by reminder time -- pure (no Application.OnTime calls, no MsgBox),
' parameterized by todayDate/nowTime rather than reading Date/Now()
' directly so it's independently testable with controlled values.
' "Reminders count down toward the meeting" (confirmed decision): the
' Nth reminder (0-based) is ReminderFirstOffset - N*ReminderInterval
' minutes before the meeting; a reminder that would land at a negative
' offset (settings where Count*Interval overshoots FirstOffset) is
' silently skipped rather than firing after the meeting has started.
' outInfos entries are "MeetingName<sep>HH:MM<sep>MinutesBefore" (see
' REMINDER_INFO_SEP), consumed by FireMeetingReminder via Split().
Public Sub CollectDueReminders(todayDate As Date, nowTime As Date, ByRef outTimes() As Date, ByRef outInfos() As String, ByRef outCount As Long)
    outCount = 0
    ReDim outTimes(1 To 1000)
    ReDim outInfos(1 To 1000)

    Dim scheduleWs As Worksheet, dataWs As Worksheet
    Set scheduleWs = ScheduleSheet()
    Set dataWs = ThisWorkbook.Sheets("Schedule Data")

    Dim firstOffset As Double, cnt As Long, interval As Double
    firstOffset = scheduleWs.Range("ReminderFirstOffset").Value
    cnt = scheduleWs.Range("ReminderCount").Value
    interval = scheduleWs.Range("ReminderInterval").Value

    Dim lastRow As Long
    lastRow = modHierarchy.LastHierarchyRow(dataWs, MT_COL_KEY, MT_FIRST_DATA_ROW)

    Dim r As Long
    For r = MT_FIRST_DATA_ROW To lastRow
        If Len(Trim$(dataWs.Cells(r, MT_COL_KEY).Value & "")) > 0 Then
            If MeetingOccursOn(dataWs, r, todayDate) Then
                Dim startTimeOfDay As Date
                startTimeOfDay = dataWs.Cells(r, MT_COL_START_TIME).Value
                Dim meetingDateTime As Date
                meetingDateTime = todayDate + startTimeOfDay

                Dim meetingName As String
                meetingName = dataWs.Cells(r, MT_COL_MEETING_NAME).Value

                Dim i As Long
                For i = 0 To cnt - 1
                    Dim minutesBefore As Double
                    minutesBefore = firstOffset - i * interval
                    If minutesBefore >= 0 Then
                        Dim reminderTime As Date
                        reminderTime = meetingDateTime - (minutesBefore / 1440#)
                        If reminderTime > nowTime Then
                            outCount = outCount + 1
                            outTimes(outCount) = reminderTime
                            outInfos(outCount) = meetingName & REMINDER_INFO_SEP & Format(startTimeOfDay, "hh:nn") & REMINDER_INFO_SEP & CStr(CLng(minutesBefore))
                        End If
                    End If
                Next i
            End If
        End If
    Next r

    ' Insertion sort ascending by time -- small N, correctness over
    ' cleverness. Required so the FIFO queue (built from this same order
    ' in ScheduleReminders) matches the order Application.OnTime actually
    ' fires in, since reminders across DIFFERENT meetings are collected
    ' in Meetings-table row order above, not chronological order.
    Dim a As Long, b As Long
    For a = 2 To outCount
        Dim tTime As Date, tInfo As String
        tTime = outTimes(a)
        tInfo = outInfos(a)
        b = a - 1
        Do While b >= 1
            If outTimes(b) > tTime Then
                outTimes(b + 1) = outTimes(b)
                outInfos(b + 1) = outInfos(b)
                b = b - 1
            Else
                Exit Do
            End If
        Loop
        outTimes(b + 1) = tTime
        outInfos(b + 1) = tInfo
    Next a
End Sub

' Production entry point -- called from Workbook_Open (a later Phase 3
' item) and Schedule Data's own Worksheet_Change (§13.5), plus directly
' from frmScheduleDay/frmReminderSettings after they write a change (which
' runs with EnableEvents off, so the Worksheet_Change hook never
' double-fires for those). Cancels every previously scheduled
' Application.OnTime call, then reschedules today's remaining reminders.
Public Sub ScheduleReminders()
    On Error GoTo ErrHandler

    CancelAllScheduledReminders

    Dim times() As Date, infos() As String, cnt As Long
    CollectDueReminders Date, Now(), times, infos, cnt

    Dim i As Long
    For i = 1 To cnt
        mReminderQueue.Add infos(i)
        mScheduledTimes.Add times(i)
        Application.OnTime EarliestTime:=times(i), Procedure:="FireMeetingReminder"
    Next i
    Exit Sub
ErrHandler:
    ' Never let an unhandled error here block Workbook_Open (which calls
    ' this unattended) or any other caller.
End Sub

' Called from Workbook_BeforeClose (a later Phase 3 item) as well as
' internally by ScheduleReminders. Canceling a time that already fired or
' was never scheduled is a safe no-op in VBA, so no existence check is
' needed beyond the Collections themselves.
Public Sub CancelAllScheduledReminders()
    If mScheduledTimes Is Nothing Then
        Set mScheduledTimes = New Collection
        Set mReminderQueue = New Collection
        Exit Sub
    End If

    Dim t As Variant
    For Each t In mScheduledTimes
        Application.OnTime EarliestTime:=CDate(t), Procedure:="FireMeetingReminder", Schedule:=False
    Next t

    Set mScheduledTimes = New Collection
    Set mReminderQueue = New Collection
End Sub

' The fixed Application.OnTime callback target -- see module header
' comment for why this pops a FIFO queue instead of receiving parameters
' directly (OnTime can't pass any).
'
' Top-level On Error is not optional here the way it is for a regular
' Application.Run-tested entry point -- this fires unattended, live, on
' the user's own screen (user-reported: it used to pop a blocking
' "Run-time error" dialog with no handler at all). Also guards against
' Application.Ready being False (OnTime firing while the user is mid-edit
' in a cell) -- showing a modal UserForm in that state can throw, so this
' reschedules itself a few seconds later instead of losing the reminder
' or crashing.
Public Sub FireMeetingReminder()
    On Error GoTo ErrHandler

    If Not Application.Ready Then
        Application.OnTime EarliestTime:=Now + TimeSerial(0, 0, 5), Procedure:="FireMeetingReminder"
        Exit Sub
    End If

    Dim info As String
    info = PopNextReminderInfo()
    If Len(info) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(info, REMINDER_INFO_SEP)
    If UBound(parts) < 2 Then Exit Sub

    frmMeetingReminder.ShowReminder parts(0), parts(1), CLng(parts(2))
    Exit Sub
ErrHandler:
    ' Never let an unhandled error surface as a blocking modal dialog on
    ' the user's screen -- worst case this reminder is silently missed,
    ' which is far better than an unattended runtime-error popup.
End Sub

' Split from FireMeetingReminder specifically so tests can verify queue
' contents/order without ever triggering frmMeetingReminder's Me.Show.
Public Function PopNextReminderInfo() As String
    If mReminderQueue Is Nothing Then Exit Function
    If mReminderQueue.Count = 0 Then Exit Function
    PopNextReminderInfo = mReminderQueue(1)
    mReminderQueue.Remove 1
End Function

Public Function ReminderQueueCount() As Long
    If mReminderQueue Is Nothing Then
        ReminderQueueCount = 0
    Else
        ReminderQueueCount = mReminderQueue.Count
    End If
End Function

Public Function ScheduledTimesCount() As Long
    If mScheduledTimes Is Nothing Then
        ScheduledTimesCount = 0
    Else
        ScheduledTimesCount = mScheduledTimes.Count
    End If
End Function
