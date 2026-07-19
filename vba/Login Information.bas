Attribute VB_Name = "Login Information"
' NOT a standard module -- this file's content is injected directly into
' the Login Information WORKSHEET's own code module by Phase 4's
' inject_vba.py (see PLAN.md §18's sheet-named-.bas convention). Login
' Information is a flat reference table (§14) with no hierarchy engine
' and no Worksheet_Change handler -- the only interactive column is
' Notes, reusing frmNotes exactly as Archive does (§19 Phase 5 correction
' note), so like Archive this sheet has no companion modXxx.bas.

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- see Tasks.bas's own note; a
    ' merged-cell Target legitimately reports Cells.Count > 1 and must
    ' not be excluded (PLAN.md §19 Phase 5).
    If Target.Row < 2 Then Exit Sub

    Select Case Target.Column
        Case 7    ' Notes
            frmNotes.ShowForCell Target
            Cancel = True
    End Select
End Sub
