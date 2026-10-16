Attribute VB_Name = "Recurring Tasks"
' NOT a standard module -- this file's content is injected directly into
' the Recurring Tasks WORKSHEET's own code module by Phase 4's
' inject_vba.py (see PLAN.md §18's sheet-named-.bas convention). Contains
' only the two event dispatchers (PLAN.md §12.2/§12.4); all actual logic
' lives in modRecurringTasks.bas so it stays independently testable via
' Application.Run. Column numbers mirror Tasks.bas's own dispatcher shape
' but at Recurring Tasks' own positions (Status prepended shifts
' everything right by one vs. Tasks, and Deadline/Timer/Time Spent don't
' exist here at all, §12.4).

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    Select Case Target.Column
        Case 5    ' Task
            modRecurringTasks.HandleTaskNameChange Target.Row
        Case 7    ' Subtask of
            modRecurringTasks.HandleSubtaskOfChange Target.Row
        Case 9    ' Linked Campaign
            modRecurringTasks.HandleLinkedCampaignChange Target.Row
    End Select
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- see Tasks.bas's own note; a
    ' merged-cell Target legitimately reports Cells.Count > 1 and must
    ' not be excluded (PLAN.md §19 Phase 5).
    If Target.Row < 2 Then Exit Sub

    Select Case Target.Column
        Case 3    ' up arrow
            modRecurringTasks.HandleSiblingReorder Target.Row, "up"
            Cancel = True
        Case 4    ' down arrow
            modRecurringTasks.HandleSiblingReorder Target.Row, "down"
            Cancel = True
        Case 5    ' Task -- only intercept double-click when this row has subtasks
            Dim numSub As String
            numSub = CStr(Target.Worksheet.Cells(Target.Row, 6).Value)
            If Len(numSub) > 0 And numSub <> "0" Then
                modRecurringTasks.HandleTaskCollapseDoubleClick Target.Row
                Cancel = True
            End If
        Case 16   ' Notes -- column 17 is the hidden RecurringKey (§20 bug
                   ' #2, same root cause as Tasks): a row with no key yet is
                   ' one RenumberHierarchy call away from being treated as a
                   ' blank/phantom row and having whatever's written here
                   ' silently discarded.
            If Len(Trim$(Target.Worksheet.Cells(Target.Row, 17).Value & "")) = 0 Then
                MsgBox "Type the task name first, then add notes.", vbInformation, "No task yet"
            Else
                frmNotes.ShowForCell Target
            End If
            Cancel = True
    End Select
End Sub
