Attribute VB_Name = "ThisWorkbook"
' Injected directly into the workbook's own special ThisWorkbook code
' module for the MANAGER build (same mechanism as vba/ThisWorkbook.bas,
' which this file replaces for that variant -- see inject_vba.py's
' per-variant STANDARD_MODULES/ThisWorkbook mapping, plan §8/§9).
'
' Identical to vba/ThisWorkbook.bas except: no Timer-left-running check
' (manager tracks no Timer at all, and modTasks -- the module that check
' calls into -- isn't injected into the manager build at all; keeping the
' call would fail to COMPILE, not just error at runtime, since a missing
' Public Function reference is a compile-time error in VBA). Adds a
' defensive call into the PIC List recompaction routine (plan §2) --
' backfills a gap that predates this fix the same way
' RefreshCampaignNameList already backfills Campaigns' own compacted list
' below.
'
' Every substantive action here already has its own error-swallowing
' entry point (RegenerateRecurringTasks, ScheduleReminders,
' RecompactPICList all have their own On Error handlers that never
' propagate), but Workbook_Open itself still gets a top-level handler too,
' as defense in depth -- an unhandled error firing the moment the file
' opens, before the user has done anything, would be an especially bad
' first impression, and .Protect/CalculateFullRebuild aren't self-
' protected the way the modXxx entry points are.

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

    ' Same backfill reasoning as RefreshCampaignNameList above, but for
    ' the PIC List's own compacted roster (plan §2/§4) -- a workbook that
    ' already had a gap left mid-list before this fix needs one
    ' recompaction on open; PIC List.bas's own Worksheet_Change keeps it
    ' current on its own after that.
    modTasksManager.RecompactPICList
    modTasksManager.RefreshPICListButtonStyling

    ' Same backfill reasoning again, for PIC color assignments (plan/user
    ' feedback: "a way for her to change the color of each PIC") -- a
    ' Fill Color edit on a PIC List swatch cell never fires
    ' Worksheet_Change (VBA has no formatting-change event at all), so a
    ' color she picked in a previous session and never manually refreshed
    ' gets applied here on open; the "Refresh PIC Colors" banner covers
    ' the same-session case.
    modTasksManager.HandleRefreshPICColorsClick

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
