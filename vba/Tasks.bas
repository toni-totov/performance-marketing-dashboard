Attribute VB_Name = "Tasks"
' NOT a standard module -- this file's content is injected directly into
' the Tasks WORKSHEET's own code module by Phase 4's inject_vba.py (a
' file named after a sheet, rather than modXxx, is the convention this
' project uses to mark sheet code-behind vs. a standard module). Contains
' only the two event dispatchers (PLAN.md §6.3); all actual logic lives
' in modTasks.bas so it stays independently testable via Application.Run.

Private Sub Worksheet_Change(ByVal Target As Range)
    modDesignSystem.NormalizeDateInput Target   ' typed 17.10.2026 -> real date on any locale
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    Select Case Target.Column
        Case 5    ' Task
            modTasks.HandleTaskNameChange Target.Row
        Case 7    ' Subtask of
            modTasks.HandleSubtaskOfChange Target.Row
        Case 9    ' Linked Campaign
            modTasks.HandleLinkedCampaignChange Target.Row
        Case 12   ' Status
            modTasks.HandleStatusChange Target.Row
    End Select
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- BeforeDoubleClick's Target is
    ' never a broader multi-cell selection, only ever a single cell or a
    ' merged cell's whole merge area (which legitimately reports
    ' Cells.Count > 1 and must not be excluded); confirmed the hard way
    ' on Calendar/Monthly Media Plan/Schedule/Dashboard's merged buttons
    ' (PLAN.md §19 Phase 5), removed from every BeforeDoubleClick
    ' dispatcher project-wide even where currently harmless.
    If Target.Row < 2 Then Exit Sub

    ' Row extender banner -- lives one row below the current allowance
    ' (modTasks.bas's RowExtenderBannerRow), not inside the normal
    ' FIRST_DATA_ROW..LAST_DATA_ROW data range, so it's checked separately
    ' and BEFORE the column dispatch below rather than as a Case in it --
    ' otherwise its anchor column would collide with whichever existing
    ' Case already owns that column number within the real data rows.
    If Target.Row = modTasks.RowExtenderBannerRow() And Target.Column = 1 Then
        modTasks.HandleRowExtenderClick
        Cancel = True
        Exit Sub
    End If

    Select Case Target.Column
        Case 2    ' up arrow
            modTasks.HandleSiblingReorder Target.Row, "up"
            Cancel = True
        Case 3    ' down arrow
            modTasks.HandleSiblingReorder Target.Row, "down"
            Cancel = True
        Case 5    ' Task -- only intercept double-click when this row has subtasks
            Dim numSub As String
            numSub = CStr(Target.Worksheet.Cells(Target.Row, 6).Value)
            If Len(numSub) > 0 And numSub <> "0" Then
                modTasks.HandleTaskCollapseDoubleClick Target.Row
                Cancel = True
            End If
        Case 11   ' Deadline -- column 17 is the hidden TaskKey (§20 bug #2):
                   ' a row with no TaskKey yet is one RenumberHierarchy call
                   ' away from being treated as a blank/phantom row and
                   ' having whatever's written here silently discarded.
            If Len(Trim$(Target.Worksheet.Cells(Target.Row, 17).Value & "")) = 0 Then
                MsgBox "Type the task name first, then set a deadline.", vbInformation, "No task yet"
            Else
                modTasks.HandleDeadlineDoubleClick Target.Row
            End If
            Cancel = True
        Case 14   ' Timer
            If Len(Trim$(Target.Worksheet.Cells(Target.Row, 17).Value & "")) = 0 Then
                MsgBox "Type the task name first, then start the timer.", vbInformation, "No task yet"
            Else
                modTasks.HandleTimerDoubleClick Target.Row
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
