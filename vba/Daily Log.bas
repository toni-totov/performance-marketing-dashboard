Attribute VB_Name = "Daily Log"
' NOT a standard module -- injected into the Daily Log WORKSHEET's own
' code module (PLAN.md §18's sheet-named-.bas convention). Thin
' dispatchers only; the logic lives in modDailyLog.bas (PLAN.md §22.6).

Private Sub Worksheet_Activate()
    ' Always fresh: anything archived since the last visit shows up.
    modDailyLog.RenderWeek
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' No "Target.Cells.Count > 1" guard -- the buttons are merged cells
    ' (Tasks.bas's own note, PLAN.md §19 Phase 5).
    Select Case Target.Row
        Case 1
            If Target.Column = 1 Then
                modDailyLog.GoPrevWeek
                Cancel = True
            ElseIf Target.Column = 4 Then
                modDailyLog.GoNextWeek
                Cancel = True
            End If
        Case 2
            If Target.Column <= 2 Then
                modDailyLog.GoThisWeek
            Else
                modDailyLog.CopyForStandup
            End If
            Cancel = True
    End Select
End Sub
