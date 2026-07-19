Attribute VB_Name = "Dashboard"
' NOT a standard module -- this file's content is injected directly into
' the Dashboard WORKSHEET's own code module by Phase 4's inject_vba.py
' (see PLAN.md §18's sheet-named-.bas convention). One button, one
' built-in call (§4) -- no companion modDashboard.bas, since there's no
' logic beyond the recalc itself to make independently testable (unlike
' Monthly Media Plan's reset button, which has real clear/advance logic
' behind it).

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- the Refresh Now button below is
    ' a merged cell, whose Target legitimately reports Cells.Count > 1;
    ' that guard used to silently swallow its double-click (confirmed via
    ' live manual testing, PLAN.md §19 Phase 5), removed project-wide.
    If Target.Row = 4 And Target.Column = 1 Then   ' Refresh Now button (merged A4:C4)
        Application.CalculateFullRebuild
        Cancel = True
    End If
End Sub
