Attribute VB_Name = "modRecurringTasksWriter"
Option Explicit

' Manager-variant half of modRecurringTasks.bas's "write a generated
' instance into Tasks" logic -- injected under the SAME logical component
' name "modRecurringTasksWriter" as modRecurringTasksWriter.bas (the
' employee version this file replaces for that variant), so
' modRecurringTasks.bas's own qualified calls resolve identically in both
' builds without ever referencing modTasks/modTasksManager directly. See
' modRecurringTasksWriter.bas's own header comment for the full
' rationale.
'
' No Timer/Time Spent at all for manager -- those two writes simply don't
' exist here. PIC and Completed Subtasks are also left untouched (stay
' blank), matching HandleTaskNameChange's own "new task" convention in
' modTasksManager.bas (blank PIC -> unassigned group; Completed Subtasks
' blank, not zero, until a real completion first increments it).

Private Const TASKS_COL_PROJECT As Long = 10
Private Const TASKS_COL_LINKED_CAMPAIGN As Long = 11
Private Const TASKS_COL_PRIORITY As Long = 12
Private Const TASKS_COL_DEADLINE As Long = 13
Private Const TASKS_COL_STATUS As Long = 14
Private Const TASKS_COL_LINKED_CAMPAIGN_KEY As Long = 21

Public Function NextGeneratedTaskKey() As Double
    NextGeneratedTaskKey = modTasksManager.NextTaskKey()
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
End Sub

' V4 skip support is employee-only (PLAN.md §22.3) -- the manager build
' has no Skipped status or RecurringSourceKey column, so these are
' deliberate no-ops with the same signatures modRecurringTasks.bas calls.
Public Sub WriteRecurringSource(tasksWs As Worksheet, destRow As Long, sourceKey As Double)
End Sub

Public Function SkipMissedInstances(sourceKey As Double) As Long
    SkipMissedInstances = 0
End Function

Public Sub FinalizeGeneration(tasksWs As Worksheet)
    RenumberHierarchy tasksWs, modTasksManager.TasksManagerColMap()
    modTasksManager.RefreshButtonStyling tasksWs
End Sub
