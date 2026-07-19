Attribute VB_Name = "Calendar"
' NOT a standard module -- this file's content is injected directly into
' the Calendar WORKSHEET's own code module by Phase 4's inject_vba.py (see
' PLAN.md §18's sheet-named-.bas convention). One dispatcher, branching by
' Row+Column rather than Column alone (§6.3's usual convention), since
' Calendar's interactive elements are one-off single cells (nav buttons,
' Confirm/Cancel) rather than per-row repeating columns. All actual logic
' lives in modCalendar.bas so it stays independently testable via
' Application.Run.

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' NOTE: no "Target.Cells.Count > 1" guard here -- Worksheet_BeforeDoubleClick's
    ' Target is never a broader multi-cell selection, only ever a single
    ' cell or (for a merged cell, like Confirm/Cancel below) its whole
    ' merge area, which legitimately reports Cells.Count > 1. That guard
    ' used to silently swallow every double-click on Confirm/Cancel,
    ' confirmed via live manual testing (PLAN.md §19 Phase 5) -- removed
    ' project-wide, not just here.

    ' Row 1: Prev / Today / Next banner buttons
    If Target.Row = 1 Then
        Select Case Target.Column
            Case 1    ' Prev
                modCalendar.GoToPrevMonth
                Cancel = True
            Case 6    ' Today
                modCalendar.GoToToday
                Cancel = True
            Case 7    ' Next
                modCalendar.GoToNextMonth
                Cancel = True
        End Select
        Exit Sub
    End If

    ' Rows 3-8, columns A-G: the day grid
    If Target.Row >= 3 And Target.Row <= 8 And Target.Column >= 1 And Target.Column <= 7 Then
        modCalendar.SelectDay Target
        Cancel = True
        Exit Sub
    End If

    ' Row 12: Confirm / Cancel
    If Target.Row = 12 Then
        Select Case Target.Column
            Case 1    ' Confirm (merged A12:C12)
                modCalendar.ConfirmSelection
                Cancel = True
            Case 5    ' Cancel (merged E12:G12)
                modCalendar.CancelSelection
                Cancel = True
        End Select
    End If
End Sub
