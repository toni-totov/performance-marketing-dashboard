Attribute VB_Name = "Archive"
' NOT a standard module -- this file's content is injected directly into
' the Archive WORKSHEET's own code module by Phase 4's inject_vba.py (see
' PLAN.md §18's sheet-named-.bas convention). Archive is a flat,
' append-only historical log (§7) -- no hierarchy engine, no collapse/
' expand, no renumbering, and no Worksheet_Change handler, since every
' cell is protected against direct typing except through VBA (§7.1's
' correction note). The only interactive column is Notes, reusing
' frmNotes exactly as Tasks does (§6.3) -- there's no Archive-specific
' logic beyond routing the double-click, so unlike Tasks this sheet has
' no companion modArchive.bas.

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- see Tasks.bas's own note; a
    ' merged-cell Target legitimately reports Cells.Count > 1 and must
    ' not be excluded (PLAN.md §19 Phase 5).
    If Target.Row < 2 Then Exit Sub

    Select Case Target.Column
        Case 13   ' Notes
            frmNotes.ShowForCell Target
            Cancel = True
    End Select
End Sub
