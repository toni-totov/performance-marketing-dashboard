Attribute VB_Name = "Schedule"
' NOT a standard module -- this file's content is injected directly into
' the Schedule WORKSHEET's own code module by Phase 4's inject_vba.py
' (see PLAN.md §18's sheet-named-.bas convention). Branches by Row+Column
' like Calendar's own dispatcher (§8), since the interactive elements are
' one-off cells (nav banner, Reminder Settings button) or a grid of date
' cells, not per-row repeating columns. All actual logic lives in
' modSchedule.bas so it stays independently testable via Application.Run.

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- the Reminder Settings button
    ' below is a merged cell, whose Target legitimately reports
    ' Cells.Count > 1; that guard used to silently swallow its double-
    ' click (confirmed via live manual testing, PLAN.md §19 Phase 5),
    ' removed project-wide.

    ' Row 1: Prev / Today / Next banner buttons
    If Target.Row = 1 Then
        Select Case Target.Column
            Case 1    ' Prev
                modSchedule.GoToPrevMonth
                Cancel = True
            Case 6    ' Today
                modSchedule.GoToToday
                Cancel = True
            Case 7    ' Next
                modSchedule.GoToNextMonth
                Cancel = True
        End Select
        Exit Sub
    End If

    ' Rows 3-14: the day grid, 2 physical rows per calendar week (a date
    ' row immediately followed by a meeting-count badge row, §13.2) --
    ' only the date rows (3, 5, 7, 9, 11, 13) are real double-click
    ' targets.
    If Target.Row >= 3 And Target.Row <= 14 And Target.Column >= 1 And Target.Column <= 7 Then
        If (Target.Row - 3) Mod 2 = 0 Then
            frmScheduleDay.ShowForDate Target.Value
            Cancel = True
        End If
        Exit Sub
    End If

    ' Row 15: Reminder Settings button (merged A15:C15)
    If Target.Row = 15 And Target.Column = 1 Then
        frmReminderSettings.ShowSettings
        Cancel = True
    End If
End Sub
