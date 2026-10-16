Attribute VB_Name = "Archive"
' NOT a standard module -- this file's content is injected directly into
' the Archive WORKSHEET's own code module for the MANAGER build (same
' sheet-code-behind convention as vba/Archive.bas, which this file
' replaces for that variant -- see inject_vba.py's per-variant
' SHEET_CODE_BEHIND mapping, plan §9). Archive is a flat, append-only
' historical log in both variants (§7) -- no hierarchy engine, no
' collapse/expand, no renumbering, and no Worksheet_Change handler, since
' every cell is protected against direct typing except through VBA (§7.1's
' correction note). The only interactive column is Notes, reusing
' frmNotes exactly as Tasks does (§6.3) -- there's no Archive-specific
' logic beyond routing the double-click, so unlike Tasks this sheet has
' no companion modArchiveManager.bas.
'
' Cannot reuse vba/Archive.bas as-is -- it hardcodes Case 13 for the Notes
' column double-click, but Notes is column 14 in the manager layout
' (sheets/archive.py, variant="manager": ID, Start Date, Task, PIC, Total
' Subtasks, Completed Subtasks, Subtask of, Project/Category, Linked
' Campaign, Priority, Deadline, Status, Days Left, Notes, ...).

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- see Tasks.bas's own note; a
    ' merged-cell Target legitimately reports Cells.Count > 1 and must
    ' not be excluded (PLAN.md §19 Phase 5).
    If Target.Row < 2 Then Exit Sub

    Select Case Target.Column
        Case 14   ' Notes
            frmNotes.ShowForCell Target
            Cancel = True
    End Select
End Sub
