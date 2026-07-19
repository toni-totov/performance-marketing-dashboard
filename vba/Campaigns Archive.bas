Attribute VB_Name = "Campaigns Archive"
' NOT a standard module -- this file's content is injected directly into
' the Campaigns Archive WORKSHEET's own code module by Phase 4's
' inject_vba.py (see PLAN.md §18's sheet-named-.bas convention). Flat,
' append-only historical log (§11) -- no hierarchy, no Gantt, no
' collapse/expand, and no Worksheet_Change handler, since every cell is
' protected against direct typing except through VBA (same pattern as
' Tasks' own Archive, §7.1/§11.2). The only interactive column is Notes,
' reusing frmNotes exactly as Archive/Tasks/Campaigns do -- no logic
' beyond routing the double-click, so this sheet has no companion
' modCampaignsArchive.bas.

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- see Tasks.bas's own note; a
    ' merged-cell Target legitimately reports Cells.Count > 1 and must
    ' not be excluded (PLAN.md §19 Phase 5).
    If Target.Row < 2 Then Exit Sub

    Select Case Target.Column
        Case 6    ' Notes
            frmNotes.ShowForCell Target
            Cancel = True
    End Select
End Sub
