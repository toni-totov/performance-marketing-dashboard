Attribute VB_Name = "modRecurringTasksWriter"
Option Explicit

' Employee-variant half of modRecurringTasks.bas's "write a generated
' instance into Tasks" logic (PLAN.md, "Manager-variant workbook" §8/§9
' gap found during Phase 8 build-orchestration work -- not present in the
' original plan text). modRecurringTasks.bas itself stays genuinely
' shared/unchanged between variants (it only ever touches the Recurring
' Tasks sheet's own column layout, which is identical either way); only
' the handful of fields that differ in Tasks' column position between
' variants -- Project/Linked Campaign/Linked Campaign Key/Priority/
' Deadline/Status, plus Timer/Time Spent (employee only) -- live here,
' referenced from modRecurringTasks.bas by the logical component name
' "modRecurringTasksWriter", injected from THIS file for the employee
' build and from modRecurringTasksWriterManager.bas for the manager build
' (inject_vba.py, see its own module-list comment) -- same "one qualified
' name, variant-selected source file" mechanism TasksManager.bas/
' ArchiveManager.bas already use for sheet code-behind, extended here to
' a standard module. This is what lets modRecurringTasks.bas reference
' "modRecurringTasksWriter.X" without ever having to know modTasks vs.
' modTasksManager exists -- referencing either of THOSE directly from a
' genuinely shared file would be a compile error in whichever variant
' doesn't have that module (the exact problem ThisWorkbook.bas's Timer
' check already had to work around, plan §8).

Private Const TASKS_COL_PROJECT As Long = 8
Private Const TASKS_COL_LINKED_CAMPAIGN As Long = 9
Private Const TASKS_COL_PRIORITY As Long = 10
Private Const TASKS_COL_DEADLINE As Long = 11
Private Const TASKS_COL_STATUS As Long = 12
Private Const TASKS_COL_TIMER As Long = 14
Private Const TASKS_COL_TIME_SPENT As Long = 15
Private Const TASKS_COL_LINKED_CAMPAIGN_KEY As Long = 22

Public Function NextGeneratedTaskKey() As Double
    NextGeneratedTaskKey = modTasks.NextTaskKey()
End Function

Public Sub WriteVariantSpecificFields(tasksWs As Worksheet, destRow As Long, projectVal As Variant, _
                                       linkedCampaignVal As Variant, linkedCampaignKeyVal As Variant, _
                                       priorityVal As Variant, deadlineVal As Date)
    tasksWs.Cells(destRow, TASKS_COL_PROJECT).Value = projectVal
    tasksWs.Cells(destRow, TASKS_COL_LINKED_CAMPAIGN).Value = linkedCampaignVal
    tasksWs.Cells(destRow, TASKS_COL_LINKED_CAMPAIGN_KEY).Value = linkedCampaignKeyVal
    tasksWs.Cells(destRow, TASKS_COL_PRIORITY).Value = priorityVal
    tasksWs.Cells(destRow, TASKS_COL_DEADLINE).Value = deadlineVal
    tasksWs.Cells(destRow, TASKS_COL_STATUS).Value = "Not Started"

    ' Timer/Time Spent defaults (user-reported, PLAN.md §19 Phase 5): a
    ' generated instance bypasses modTasks.HandleTaskNameChange entirely,
    ' so it never picked up these "new task" defaults without this.
    tasksWs.Cells(destRow, TASKS_COL_TIMER).Value = ICON_TIMER_RUNNING & " Start"
    ' 0, not blank -- Time Spent (hrs) is hidden, but its value drives the
    ' visible friendly "Time Spent" column's formula (=IF(ts="","",...)),
    ' so a genuinely blank hrs cell showed as a blank friendly cell
    ' instead of "0h 00m".
    tasksWs.Cells(destRow, TASKS_COL_TIME_SPENT).Value = 0
End Sub

' V4 (PLAN.md §22.4) -- stamps a generated ROOT row with its template's
' RecurringKey, so "If Missed = Skip" can find it again later.
Public Sub WriteRecurringSource(tasksWs As Worksheet, destRow As Long, sourceKey As Double)
    modTasks.WriteRecurringSourceKey destRow, sourceKey
End Sub

' V4 (PLAN.md §22.4) -- archives a template's missed open instances as
' Skipped. Returns how many were skipped.
Public Function SkipMissedInstances(sourceKey As Double) As Long
    SkipMissedInstances = modTasks.SkipMissedRecurringInstances(sourceKey)
End Function

Public Sub FinalizeGeneration(tasksWs As Worksheet)
    RenumberHierarchy tasksWs, modTasks.TasksColMap()
    modTasks.RefreshButtonStyling tasksWs
End Sub
