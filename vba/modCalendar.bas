Attribute VB_Name = "modCalendar"
Option Explicit

' Calendar sheet business logic (PLAN.md §8.6). Every Public
' Function/Sub here is independently callable via Application.Run for
' testing. No hierarchy/renumbering engine is involved -- this sheet has
' no tree structure, just navigation state (SelectedYear/SelectedMonth),
' a picked date/time (SelectedDate/SelectedHour/SelectedMinute), and the
' generic 4-value handoff contract (modCalendarHandoff.bas, Phase 1).
'
' Confirm's guard checks (no date selected / no pending handoff / target
' not found) each show a MsgBox on their production path, which would
' block headless COM testing -- so the actual write+navigate action is
' split into the Public ApplyConfirmedSelection, testable directly and
' bypassing ConfirmSelection's MsgBox branches, same reasoning as Tasks'
' CascadeDelete/WouldCreateCycle split (PLAN.md §19 Phase 3).

Private Function CalendarSheet() As Worksheet
    Set CalendarSheet = ThisWorkbook.Sheets("Calendar")
End Function

Public Sub GoToPrevMonth()
    Dim ws As Worksheet
    Set ws = CalendarSheet()
    Dim y As Long, m As Long
    y = ws.Range("SelectedYear").Value
    m = ws.Range("SelectedMonth").Value
    m = m - 1
    If m < 1 Then
        m = 12
        y = y - 1
    End If
    ws.Range("SelectedYear").Value = y
    ws.Range("SelectedMonth").Value = m
End Sub

Public Sub GoToNextMonth()
    Dim ws As Worksheet
    Set ws = CalendarSheet()
    Dim y As Long, m As Long
    y = ws.Range("SelectedYear").Value
    m = ws.Range("SelectedMonth").Value
    m = m + 1
    If m > 12 Then
        m = 1
        y = y + 1
    End If
    ws.Range("SelectedYear").Value = y
    ws.Range("SelectedMonth").Value = m
End Sub

Public Sub GoToToday()
    Dim ws As Worksheet
    Set ws = CalendarSheet()
    ws.Range("SelectedYear").Value = Year(Date)
    ws.Range("SelectedMonth").Value = Month(Date)
End Sub

' dayCell is one of the 42 day-grid cells (§8.2) -- each holds a real
' DATE-arithmetic formula (week_grid.py), so .Value is a genuine date
' serial number, never text (Bulgarian-locale-safe by construction, per
' CLAUDE.md's native-COM-type rule).
Public Sub SelectDay(dayCell As Range)
    CalendarSheet().Range("SelectedDate").Value = dayCell.Value
End Sub

' The actual write+navigate action, assuming preconditions (a date is
' selected, a real target was resolved) are already met -- see module
' docstring for why this is split from ConfirmSelection.
Public Sub ApplyConfirmedSelection(targetCell As Range, selectedDateTime As Date)
    targetCell.Value = selectedDateTime
    modCalendarHandoff.ClearCalendarHandoff
    ReturnToCell targetCell
End Sub

' Back to the cell the picker was opened from, showing its whole row
' (PLAN.md §22.12, user-reported). This used to be Application.Goto
' target, Scroll:=True, which makes the target the TOP-LEFT visible cell
' -- for Tasks' Deadline (column K) that scrolled ID and Task off-screen,
' so the user had to scroll back every time. Now: scroll fully left, put
' the row a few lines below the top, and select the target cell. Display
' only, so a failure here never undoes the value already written.
Public Sub ReturnToCell(targetCell As Range)
    targetCell.Worksheet.Activate
    targetCell.Select
    On Error Resume Next
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = Application.WorksheetFunction.Max(1, targetCell.Row - 5)
End Sub

Public Sub ConfirmSelection()
    Dim ws As Worksheet
    Set ws = CalendarSheet()

    If Len(Trim$(ws.Range("SelectedDate").Value & "")) = 0 Then
        MsgBox "Please select a date first.", vbExclamation, "No Date Selected"
        Exit Sub
    End If

    If Not modCalendarHandoff.HasPendingCalendarHandoff() Then
        MsgBox "No pending date selection -- Calendar was opened for browsing only.", vbInformation, "Nothing to Confirm"
        Exit Sub
    End If

    Dim targetCell As Range
    Set targetCell = modCalendarHandoff.ResolveCalendarHandoffTarget()
    If targetCell Is Nothing Then
        MsgBox "Could not find the original task/row to update. The selection was not applied.", vbExclamation, "Handoff Target Not Found"
        modCalendarHandoff.ClearCalendarHandoff
        Exit Sub
    End If

    ApplyConfirmedSelection targetCell, ws.Range("SelectedDateTime").Value
End Sub

' No MsgBox on any path -- directly testable via Application.Run.
Public Sub CancelSelection()
    Dim targetCell As Range
    Set targetCell = Nothing
    If modCalendarHandoff.HasPendingCalendarHandoff() Then
        Set targetCell = modCalendarHandoff.ResolveCalendarHandoffTarget()
    End If

    modCalendarHandoff.ClearCalendarHandoff

    If Not targetCell Is Nothing Then
        ReturnToCell targetCell
    End If
End Sub
