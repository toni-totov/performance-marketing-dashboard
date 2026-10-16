Attribute VB_Name = "PIC List"
' NOT a standard module -- sheet code-behind for the "PIC List" worksheet
' (manager build only, see sheets/pic_list.py). Recompaction and reorder
' logic live in modTasksManager.bas so they stay independently testable
' via Application.Run (RecompactPICList/HandlePICListReorder).

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Row < 2 Then Exit Sub
    If Target.Column <> 1 Then Exit Sub

    ' Deliberately no "Target.Cells.Count > 1" guard here (unlike Tasks'
    ' dispatcher) -- RecompactPICList always re-scans the whole column
    ' regardless of which cells changed, so a multi-cell paste/clear is
    ' just as valid a trigger as a single-cell edit, and is in fact the
    ' more likely way to actually create a gap (selecting several names
    ' and pressing Delete at once).
    modTasksManager.RecompactPICList

    ' A name being typed or cleared changes which rows need the Up/Down
    ' button look -- refresh after every edit, same "re-scan and
    ' reconcile" reasoning as RecompactPICList itself (PLAN.md, "Manager-
    ' variant workbook" user feedback after using the manager build).
    modTasksManager.RefreshPICListButtonStyling
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- see Tasks.bas's own note; a
    ' merged-cell Target legitimately reports Cells.Count > 1 and must
    ' not be excluded (PLAN.md §19 Phase 5).
    If Target.Row < 2 Then Exit Sub

    ' "Refresh PIC Colors" banner -- fixed row (this list never grows, so
    ' unlike Tasks' row-extender banner there's no moving-target position
    ' to compute), checked before the column dispatch below, same
    ' reasoning as Tasks.bas's own banner check.
    If Target.Row = 202 And Target.Column = 1 Then
        modTasksManager.HandleRefreshPICColorsClick
        Cancel = True
        Exit Sub
    End If

    Select Case Target.Column
        Case 2    ' up arrow
            modTasksManager.HandlePICListReorder Target.Row, "up"
            Cancel = True
        Case 3    ' down arrow
            modTasksManager.HandlePICListReorder Target.Row, "down"
            Cancel = True
    End Select
End Sub
