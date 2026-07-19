Attribute VB_Name = "ThisWorkbook"
' Injected directly into the workbook's own special ThisWorkbook code
' module by Phase 4's inject_vba.py -- the same "inject into an existing
' component" mechanism as sheet-named .bas files (PLAN.md §18), just
' targeting the one always-present ThisWorkbook component (whose VBA
' component name is always literally "ThisWorkbook", unlike a regular
' sheet's auto-assigned CodeName, so no lookup is needed to find it).
'
' Every substantive action here already has its own error-swallowing
' entry point (RegenerateRecurringTasks, ScheduleReminders,
' FindRunningTimerTaskName all have their own On Error handlers that
' never propagate), but Workbook_Open itself still gets a top-level
' handler too, as defense in depth -- an unhandled error firing the
' moment the file opens, before the user has done anything, would be an
' especially bad first impression, and .Protect/CalculateFullRebuild
' aren't self-protected the way the modXxx entry points are.

Private Sub Workbook_Open()
    On Error GoTo ErrHandler

    Application.CalculateFullRebuild

    ' UserInterfaceOnly does not persist across file close/reopen
    ' (PLAN.md §7.1/§11.1 correction notes) -- must be re-asserted here,
    ' or the Tasks/Campaigns archiving cascades' own VBA writes would be
    ' blocked by the sheets' base protection.
    ThisWorkbook.Sheets("Archive").Protect UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Campaigns Archive").Protect UserInterfaceOnly:=True

    modRecurringTasks.RegenerateRecurringTasks
    modSchedule.ScheduleReminders

    ' Backfills the Linked Campaign dropdown's compacted source list
    ' (PLAN.md §19 Phase 5, sheets/campaigns.py module docstring) -- it's
    ' VBA-maintained, not formula-driven, so a workbook that already had
    ' campaigns typed in before this fix needs one rebuild on open;
    ' every edit after that keeps it current on its own.
    modCampaigns.RefreshCampaignNameList

    Dim runningTask As String
    runningTask = modTasks.FindRunningTimerTaskName()
    If Len(runningTask) > 0 Then
        MsgBox "The timer for '" & runningTask & "' was still running when this file was last closed.", vbExclamation, "Timer Left Running"
    End If

    Exit Sub
ErrHandler:
    ' Never let an unhandled error block the file from opening.
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modSchedule.CancelAllScheduledReminders
    ' Belt-and-suspenders restore -- see Workbook_Deactivate's own note.
    ' Closing the LAST open workbook doesn't fire Deactivate first on
    ' every Excel version, so this is the backstop for that case.
    Application.DisplayFormulaBar = True
End Sub

' Modernization pass (§17.10, PLAN.md §19 Phase 6, user-requested):
' Application.DisplayFormulaBar is a genuine Excel-wide setting, not
' saved per-workbook -- toggling it in Workbook_Open alone would leak
' into every OTHER workbook the user has open in the same Excel session.
' Scoped instead to Activate/Deactivate so the formula bar only
' disappears while this workbook is the active window, and comes right
' back the moment focus moves to any other workbook.
Private Sub Workbook_Activate()
    On Error Resume Next
    Application.DisplayFormulaBar = False
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    Application.DisplayFormulaBar = True
End Sub
