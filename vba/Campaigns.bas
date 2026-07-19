Attribute VB_Name = "Campaigns"
' NOT a standard module -- this file's content is injected directly into
' the Campaigns WORKSHEET's own code module by Phase 4's inject_vba.py
' (see PLAN.md §18's sheet-named-.bas convention). All actual logic lives
' in modCampaigns.bas so it stays independently testable via
' Application.Run. Notes reuses frmNotes directly (same as Tasks/Archive),
' no modCampaigns wrapper needed for it.

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 3 Then Exit Sub

    Select Case Target.Column
        Case 3    ' Campaign Name
            modCampaigns.HandleCampaignNameChange Target.Row
        Case 7    ' Status
            modCampaigns.HandleStatusChange Target.Row
    End Select
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- see Tasks.bas's own note; a
    ' merged-cell Target legitimately reports Cells.Count > 1 and must
    ' not be excluded (PLAN.md §19 Phase 5).
    If Target.Row < 3 Then Exit Sub

    Select Case Target.Column
        Case 1    ' up arrow (§19 Phase 5, user-requested)
            modCampaigns.HandleCampaignReorder Target.Row, "up"
            Cancel = True
        Case 2    ' down arrow
            modCampaigns.HandleCampaignReorder Target.Row, "down"
            Cancel = True
        Case 3    ' Campaign Name -- only intercept when this slot has linked tasks
            Dim numTasks As String
            numTasks = CStr(Target.Worksheet.Cells(Target.Row, 4).Value)
            If Len(numTasks) > 0 And numTasks <> "0" Then
                modCampaigns.HandleCampaignDoubleClick Target.Row
                Cancel = True
            End If
        Case 9    ' Notes
            frmNotes.ShowForCell Target
            Cancel = True
    End Select
End Sub
