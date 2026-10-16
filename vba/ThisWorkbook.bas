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

    ' One-time corrective assertion, NOT a per-activation toggle -- see
    ' this module's own trailing comment for why those are different and
    ' why only the latter was ever the actual problem. Fixes the formula
    ' bar back on if some OTHER workbook (e.g. an older file predating
    ' this fix) left the Excel-wide DisplayFormulaBar setting off --
    ' that setting persists across files/sessions on a machine, so this
    ' file's own window inherits whatever it was last left as otherwise.
    Application.DisplayFormulaBar = True

    Application.CalculateFullRebuild

    ' UserInterfaceOnly does not persist across file close/reopen
    ' (PLAN.md §7.1/§11.1 correction notes) -- must be re-asserted here,
    ' or the Tasks/Campaigns archiving cascades' own VBA writes would be
    ' blocked by the sheets' base protection.
    ThisWorkbook.Sheets("Archive").Protect UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Campaigns Archive").Protect UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Daily Log").Protect UserInterfaceOnly:=True

    modRecurringTasks.RegenerateRecurringTasks
    modSchedule.ScheduleReminders

    ' Backfills the Linked Campaign dropdown's compacted source list
    ' (PLAN.md §19 Phase 5, sheets/campaigns.py module docstring) -- it's
    ' VBA-maintained, not formula-driven, so a workbook that already had
    ' campaigns typed in before this fix needs one rebuild on open;
    ' every edit after that keeps it current on its own.
    modCampaigns.RefreshCampaignNameList

    ' Calendar and Schedule always open on the current month rather than
    ' whichever month was on screen when the file was last saved
    ' (PLAN.md §22.2 #3). Display-only -- a pending Calendar handoff is
    ' unaffected.
    modCalendar.GoToToday
    modSchedule.GoToToday
    modDailyLog.GoThisWeek   ' V4 (PLAN.md §22.6): open on the current week
    modToday.RenderTodayPanel   ' V4 (PLAN.md §22.7)

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
End Sub

' The §17.10 modernization pass (PLAN.md §19 Phase 6) used to hide
' Application.DisplayFormulaBar while this workbook was the active window
' (via Workbook_Activate/Deactivate) and restore it on Deactivate/
' BeforeClose. Reverted (user-reported): toggling this Excel-wide setting
' on window activation was cancelling the OS clipboard's copy/"marching
' ants" state when switching from another workbook back into this one,
' breaking copy/paste between them. Formula bar now stays on permanently,
' like any other workbook.
'
' **Follow-up (user-reported again, this time surfacing on the manager
' build):** removing the toggle made this file passive, not corrective --
' if some OTHER open workbook (e.g. an older file that predates this
' revert, still carrying the original buggy Activate/Deactivate code)
' turns the Excel-wide setting off, this file never turned it back on,
' since DisplayFormulaBar isn't a per-file setting, it's a persistent
' Excel-wide one. Workbook_Open's one-time `Application.DisplayFormulaBar
' = True` above fixes that without reintroducing the original bug -- the
' clipboard problem was specifically about toggling on EVERY activate/
' deactivate cycle (repeated, as you switch windows back and forth); a
' single assertion when the file first opens can't reproduce that, since
' it never fires again on subsequent window switches.
