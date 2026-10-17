Attribute VB_Name = "Monthly Media Plan"
' NOT a standard module -- this file's content is injected directly into
' the Monthly Media Plan WORKSHEET's own code module by Phase 4's
' inject_vba.py (see PLAN.md §18's sheet-named-.bas convention). Only two
' interactive elements, both one-off single cells (§9.4/§9.5) -- branches
' by Row+Column, same convention as Calendar's own dispatcher. The
' Worksheet_Change handler only normalizes typed day-first dates
' (modDesignSystem.NormalizeDateInput); every other column is a native
' Table dropdown/free-typed cell or the self-healing Budget (USD) formula.

Private Sub Worksheet_Change(ByVal Target As Range)
    modDesignSystem.NormalizeDateInput Target   ' typed 17.10.2026 -> real date on any locale
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- both buttons below are merged
    ' cells, whose Target legitimately reports Cells.Count > 1; that guard
    ' used to silently swallow every double-click on them (confirmed via
    ' live manual testing, PLAN.md §19 Phase 5), removed project-wide.
    If Target.Row <> 4 Then Exit Sub

    Select Case Target.Column
        Case 1    ' Set Exchange Rate button (merged A4:B4)
            frmExchangeRate.ShowRate
            Cancel = True
        Case 7    ' Start New Month button (merged G4:H4)
            modMonthlyMediaPlan.ResetMonthlyPlan
            Cancel = True
    End Select
End Sub
