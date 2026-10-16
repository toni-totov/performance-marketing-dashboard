Attribute VB_Name = "TasksManager"
' NOT a standard module -- this file's content is injected directly into
' the Tasks WORKSHEET's own code module for the MANAGER build (same
' sheet-code-behind convention as vba/Tasks.bas, which this file replaces
' for that variant -- see inject_vba.py's per-variant SHEET_CODE_BEHIND
' mapping, plan §9). Contains only the two event dispatchers; all actual
' logic lives in modTasksManager.bas so it stays independently testable
' via Application.Run.

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    Select Case Target.Column
        Case 5    ' Task
            modTasksManager.HandleTaskNameChange Target.Row
        Case 6    ' PIC
            modTasksManager.HandlePICChange Target.Row
        Case 9    ' Subtask of
            modTasksManager.HandleSubtaskOfChange Target.Row
        Case 11   ' Linked Campaign
            modTasksManager.HandleLinkedCampaignChange Target.Row
        Case 14   ' Status
            modTasksManager.HandleStatusChange Target.Row
    End Select
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- BeforeDoubleClick's Target is
    ' never a broader multi-cell selection, only ever a single cell or a
    ' merged cell's whole merge area (see vba/Tasks.bas's identical note,
    ' PLAN.md §19 Phase 5 -- the exact bug this avoids).

    If Target.Row < 2 Then Exit Sub

    ' Row extender banner -- checked before the column dispatch below,
    ' same reasoning as vba/Tasks.bas's own dispatcher.
    If Target.Row = modTasksManager.RowExtenderBannerRow() And Target.Column = 1 Then
        modTasksManager.HandleRowExtenderClick
        Cancel = True
        Exit Sub
    End If

    Select Case Target.Column
        Case 2    ' up arrow
            modTasksManager.HandleSiblingReorder Target.Row, "up"
            Cancel = True
        Case 3    ' down arrow
            modTasksManager.HandleSiblingReorder Target.Row, "down"
            Cancel = True
        Case 5    ' Task -- only intercept double-click when this row has subtasks
            Dim numSub As String
            numSub = CStr(Target.Worksheet.Cells(Target.Row, 7).Value)   ' Total Subtasks
            If Len(numSub) > 0 And numSub <> "0" Then
                modTasksManager.HandleTaskCollapseDoubleClick Target.Row
                Cancel = True
            End If
        Case 13   ' Deadline -- column 17 is the hidden TaskKey
            If Len(Trim$(Target.Worksheet.Cells(Target.Row, 17).Value & "")) = 0 Then
                MsgBox "Type the task name first, then set a deadline.", vbInformation, "No task yet"
            Else
                modTasksManager.HandleDeadlineDoubleClick Target.Row
            End If
            Cancel = True
        Case 16   ' Notes
            If Len(Trim$(Target.Worksheet.Cells(Target.Row, 17).Value & "")) = 0 Then
                MsgBox "Type the task name first, then add notes.", vbInformation, "No task yet"
            Else
                frmNotes.ShowForCell Target
            End If
            Cancel = True
    End Select
End Sub
