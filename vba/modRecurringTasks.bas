Attribute VB_Name = "modRecurringTasks"
Option Explicit

' Recurring Tasks sheet business logic (PLAN.md §12). Two halves:
'   1. A hierarchy-engine mirror of modTasks.bas (§12.2) -- Recurring
'      Tasks gets its own independent tree (own RecurringKey/ID sequence,
'      unrelated to Tasks' own), so this duplicates modTasks.bas's own
'      WouldCreateCycle/FindRowByID/FindRowByTaskKey/HandleSiblingReorder/
'      HandleTaskCollapseDoubleClick shape rather than sharing instances
'      of them (they're hardcoded to Tasks' own column layout) -- but
'      still reuses modHierarchy.bas's generic RenumberHierarchy/
'      NextSiblingOrder/LastHierarchyRow directly, since those ARE
'      parameterized by column map and carry no sheet-specific
'      assumptions.
'   2. RegenerateRecurringTasks() (§12.1) -- the watermark-based
'      generation macro, entirely new logic with no Tasks equivalent.
'      ProcessTemplate is Public (not Private) specifically so the actual
'      generation logic is testable directly via Application.Run,
'      bypassing RegenerateRecurringTasks' safety-cap MsgBox warning --
'      same reasoning as Tasks' CascadeDelete/WouldCreateCycle split.

' --- Recurring Tasks column layout (sheets/recurring_tasks.py) -------------
Private Const RT_COL_STATUS As Long = 1
Private Const RT_COL_ID As Long = 2
Private Const RT_COL_UP As Long = 3
Private Const RT_COL_DOWN As Long = 4
Private Const RT_COL_TASK As Long = 5
Private Const RT_COL_NUM_SUBTASKS As Long = 6
Private Const RT_COL_SUBTASK_OF As Long = 7
Private Const RT_COL_PROJECT As Long = 8
Private Const RT_COL_LINKED_CAMPAIGN As Long = 9
Private Const RT_COL_PRIORITY As Long = 10
Private Const RT_COL_RECURRENCE As Long = 11
Private Const RT_COL_TASK_START_ON As Long = 12
Private Const RT_COL_TIME As Long = 13            ' V4, PLAN.md §22.3
Private Const RT_COL_IF_MISSED As Long = 14       ' V4, PLAN.md §22.4 (root rows only)
Private Const RT_COL_DEADLINE_OFFSET As Long = 15
Private Const RT_COL_NOTES As Long = 16
Private Const RT_COL_KEY As Long = 17             ' RecurringKey
Private Const RT_COL_PARENT_KEY As Long = 18       ' RecurringParentKey
Private Const RT_COL_LEVEL As Long = 19
Private Const RT_COL_SIBLING_ORDER As Long = 20
Private Const RT_COL_LINKED_CAMPAIGN_KEY As Long = 21
Private Const RT_COL_LAST_GENERATED As Long = 22   ' LastGeneratedThroughDate, root rows only
Private Const RT_LAST_COL As Long = 22

Private Const RT_FIRST_DATA_ROW As Long = 2

' --- Campaigns column layout (mirrors modTasks.bas's own constants --
' see that module's header comment for why these aren't imported from a
' shared source) ---
Private Const CAMPAIGNS_FIRST_MAIN_ROW As Long = 3
Private Const CAMPAIGNS_SLOT_SIZE As Long = 10
Private Const CAMPAIGNS_NUM_SLOTS As Long = 20
Private Const CAMPAIGNS_COL_NAME As Long = 3
Private Const CAMPAIGNS_COL_KEY As Long = 25
Private Const CAMPAIGNS_ARCHIVE_COL_NAME As Long = 1
Private Const CAMPAIGNS_ARCHIVE_COL_KEY As Long = 8

' --- Tasks' own column layout (destination of generated instances --
' mirrors modTasks.bas's own Private constants; see above for why this
' isn't shared directly). Only the columns that sit at the SAME position
' in both the employee and manager Tasks layouts stay here -- by
' construction (tasks.py's own docstring: manager's PIC + Completed
' Subtasks insertions exactly offset Timer + Time Spent's removal, so
' both variants' hidden block starts at the same column Q), that's Up/
' Down/Start Date/Task/Notes and the whole hidden hierarchy block
' (TaskKey/ParentKey/Level/SiblingOrder). Project/Linked Campaign/Linked
' Campaign Key/Priority/Deadline/Status genuinely differ between variants
' (and Timer/Time Spent don't exist for manager at all), so writing those
' -- plus generating the new TaskKey itself and the post-loop
' RenumberHierarchy/RefreshButtonStyling calls, which need the right
' variant's own TasksColMap -- is delegated to modRecurringTasksWriter, a
' logical module name injected from a different source file per variant
' (modRecurringTasksWriter.bas / modRecurringTasksWriterManager.bas, see
' inject_vba.py) specifically so this file can reference it by a single
' qualified name that resolves for both builds, while modTasks/
' modTasksManager themselves (which only one variant ever has) are never
' referenced here at all -- avoiding the compile error a variant-specific
' cross-module reference would otherwise cause in the OTHER build (the
' same class of problem ThisWorkbook.bas's Timer check has, plan §8).
Private Const TASKS_COL_UP As Long = 2
Private Const TASKS_COL_DOWN As Long = 3
Private Const TASKS_COL_START_DATE As Long = 4
Private Const TASKS_COL_TASK As Long = 5
Private Const TASKS_COL_NOTES As Long = 16
Private Const TASKS_COL_TASK_KEY As Long = 17
Private Const TASKS_COL_PARENT_KEY As Long = 18
Private Const TASKS_COL_LEVEL As Long = 19
Private Const TASKS_COL_SIBLING_ORDER As Long = 20
Private Const TASKS_FIRST_DATA_ROW As Long = 2

Private Const MAX_OCCURRENCES_PER_RUN As Long = 30   ' tunable, PLAN.md §5/§12.1

' Public so modMigration can reuse the one real column map instead of
' keeping its own hard-coded copy (which silently went stale the moment
' V4 inserted Time/If Missed, PLAN.md §22.9).
Public Function RTColMap() As HierarchyColumnMap
    Dim m As HierarchyColumnMap
    m.FirstDataRow = RT_FIRST_DATA_ROW
    m.LastCol = RT_LAST_COL
    m.ColID = RT_COL_ID
    m.ColTask = RT_COL_TASK
    m.ColNumSubtasks = RT_COL_NUM_SUBTASKS
    m.ColSubtaskOf = RT_COL_SUBTASK_OF
    m.ColKey = RT_COL_KEY
    m.ColParentKey = RT_COL_PARENT_KEY
    m.ColLevel = RT_COL_LEVEL
    m.ColSiblingOrder = RT_COL_SIBLING_ORDER
    RTColMap = m
End Function

Private Function RTSheet() As Worksheet
    Set RTSheet = ThisWorkbook.Sheets("Recurring Tasks")
End Function

' ============================================================================
' Hierarchy engine (mirrors modTasks.bas's own §6.3 handlers)
' ============================================================================

' Bootstrap point for a genuinely NEW template row, same reasoning as
' Tasks' own HandleTaskNameChange -- but also initializes
' LastGeneratedThroughDate = TODAY() here. This is a technical necessity
' beyond §12.1's literal text (which never states what a brand-new
' template's initial watermark should be): treating a blank watermark as
' "generate every occurrence since the dawn of time" would immediately
' backfill up to the 30-occurrence safety cap the moment a new template
' is created, which is clearly not the intended experience. Every row is,
' by construction, root-level at the instant it's first named (Subtask of
' is only ever set afterward, in a separate edit), so this is always the
' correct moment to seed it.
Public Function HandleTaskNameChange(r As Long) As String
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = RTSheet()

    If Len(Trim$(ws.Cells(r, RT_COL_TASK).Value & "")) > 0 Then
        ' User-requested default (PLAN.md §19 Phase 5): a brand-new
        ' template starts life "Active" rather than sitting blank until
        ' the user manually picks a Status -- RegenerateRecurringTasks
        ' only ever processes rows whose Status is exactly "Active", so a
        ' blank default silently meant a newly-created template never
        ' generated anything at all.
        If Len(Trim$(ws.Cells(r, RT_COL_STATUS).Value & "")) = 0 Then
            ws.Cells(r, RT_COL_STATUS).Value = "Active"
        End If

        ' User-requested default (PLAN.md §19 Phase 5): a brand-new
        ' template starts life Linked to "None" rather than sitting
        ' blank, same fix as Tasks' own HandleTaskNameChange -- "None" is
        ' the permanent literal first entry in the Linked Campaign
        ' dropdown's source list (Campaigns!Z3), not a real campaign, so
        ' no downstream lookup (HandleLinkedCampaignChange) ever matches
        ' it. This was also silently leaving every GENERATED instance's
        ' own Linked Campaign blank too, since GenerateInstance's "None"
        ' fallback only triggers when the template's own field is blank
        ' -- fixing it here at the template level is the more direct fix.
        If Len(Trim$(ws.Cells(r, RT_COL_LINKED_CAMPAIGN).Value & "")) = 0 Then
            ws.Cells(r, RT_COL_LINKED_CAMPAIGN).Value = "None"
        End If

        If Len(Trim$(ws.Cells(r, RT_COL_KEY).Value & "")) = 0 Then
            ws.Cells(r, RT_COL_KEY).Value = NextRecurringKey()
            ' Date - 1 (yesterday), NOT Date (today) -- a second, found-
            ' live bug (PLAN.md §19 Phase 5): ProcessTemplate always
            ' computes fromDate = watermark + 1, and ComputeScheduledDates
            ' exits immediately whenever fromDate > throughDate. Seeding
            ' the watermark to TODAY made the very first generation run
            ' compute fromDate = tomorrow, always past throughDate = today
            ' -- so even a correctly "Active" brand-new template generated
            ' nothing on its first reopen. Seeding to yesterday makes the
            ' first run correctly include today.
            ws.Cells(r, RT_COL_LAST_GENERATED).Value = Date - 1

            ' User-requested default (PLAN.md §19 Phase 5): populate the
            ' Up/Down cells as soon as a template is created, same fix as
            ' Tasks' own HandleTaskNameChange -- the reorder functionality
            ' (HandleSiblingReorder, wired in "Recurring Tasks.bas") was
            ' already fully working, but with no visible icon in those
            ' cells there was no affordance that they were clickable at
            ' all.
            ws.Cells(r, RT_COL_UP).Value = ICON_UP
            ws.Cells(r, RT_COL_DOWN).Value = ICON_DOWN

            Dim colMap As HierarchyColumnMap
            colMap = RTColMap()
            ws.Cells(r, RT_COL_SIBLING_ORDER).Value = NextSiblingOrder(ws, colMap, "")
            RenumberHierarchy ws, RTColMap()
            RefreshButtonStyling ws
            RefreshLinkedCampaignLabels ws
        End If
    End If

    Application.EnableEvents = prevEvents
    HandleTaskNameChange = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandleTaskNameChange = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Re-applies the Up/Down "button" look (§19 Phase 6, user-reported) to
' every row that currently holds a real template, after any operation
' that can add/remove/reorder rows -- same reasoning as Tasks' own
' RefreshButtonStyling (modTasks.bas). No Timer column here, unlike
' Tasks -- recurring templates aren't individually time-tracked.
Public Sub RefreshButtonStyling(ws As Worksheet)
    Dim lastRow As Long, r As Long
    lastRow = LastHierarchyRow(ws, RT_COL_KEY, RT_FIRST_DATA_ROW)
    For r = RT_FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, RT_COL_KEY).Value & "")) > 0 Then
            modDesignSystem.ApplyButtonLook ws.Cells(r, RT_COL_UP)
            modDesignSystem.ApplyButtonLook ws.Cells(r, RT_COL_DOWN)
            modDesignSystem.HideNumberAsTextFlag ws.Cells(r, RT_COL_ID)
            modDesignSystem.HideNumberAsTextFlag ws.Cells(r, RT_COL_SUBTASK_OF)
        End If
    Next r
End Sub

' No Archive counterpart to check -- unlike Tasks, a recurring template is
' never archived (it's toggled Inactive instead), so only Recurring
' Tasks' own key column needs scanning.
Private Function NextRecurringKey() As Double
    Dim ws As Worksheet
    Set ws = RTSheet()

    Dim maxKey As Double
    maxKey = 0

    Dim lastRow As Long, r As Long
    lastRow = LastHierarchyRow(ws, RT_COL_KEY, RT_FIRST_DATA_ROW)
    For r = RT_FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, RT_COL_KEY).Value & "")) > 0 Then
            If CDbl(ws.Cells(r, RT_COL_KEY).Value) > maxKey Then maxKey = CDbl(ws.Cells(r, RT_COL_KEY).Value)
        End If
    Next r

    NextRecurringKey = maxKey + 1
End Function

Public Function HandleSubtaskOfChange(r As Long) As String
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = RTSheet()

    Dim typedID As String
    typedID = Trim$(ws.Cells(r, RT_COL_SUBTASK_OF).Value & "")

    Dim movingKey As Double
    movingKey = CDbl(ws.Cells(r, RT_COL_KEY).Value)

    Dim newParentKey As String
    newParentKey = ""
    If Len(typedID) > 0 Then
        Dim foundRow As Long
        foundRow = FindRowByID(ws, typedID)
        If foundRow > 0 And foundRow <> r Then
            newParentKey = CStr(ws.Cells(foundRow, RT_COL_KEY).Value)
        End If
    End If

    If Len(newParentKey) > 0 Then
        If WouldCreateCycle(ws, movingKey, newParentKey) Then
            Application.EnableEvents = prevEvents
            MsgBox "That would make this template a subtask of itself (directly or indirectly). Change not applied.", vbExclamation, "Invalid Subtask Of"
            Application.EnableEvents = False
            RenumberHierarchy ws, RTColMap()
            RefreshButtonStyling ws
            Application.EnableEvents = prevEvents
            HandleSubtaskOfChange = "REJECTED_CYCLE"
            Exit Function
        End If
    End If

    Dim currentParentKey As String
    currentParentKey = Trim$(ws.Cells(r, RT_COL_PARENT_KEY).Value & "")

    If newParentKey <> currentParentKey Then
        ws.Cells(r, RT_COL_PARENT_KEY).Value = newParentKey
        Dim colMap As HierarchyColumnMap
        colMap = RTColMap()
        ws.Cells(r, RT_COL_SIBLING_ORDER).Value = NextSiblingOrder(ws, colMap, newParentKey)
    End If

    RenumberHierarchy ws, RTColMap()
    RefreshButtonStyling ws
    RefreshLinkedCampaignLabels ws

    Application.EnableEvents = prevEvents
    HandleSubtaskOfChange = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandleSubtaskOfChange = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function HandleLinkedCampaignChange(r As Long) As String
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = RTSheet()

    Dim typedName As String
    typedName = Trim$(ws.Cells(r, RT_COL_LINKED_CAMPAIGN).Value & "")

    If Len(typedName) = 0 Then
        ws.Cells(r, RT_COL_LINKED_CAMPAIGN_KEY).Value = ""
    Else
        Dim campaignsWs As Worksheet
        Set campaignsWs = ThisWorkbook.Sheets("Campaigns")
        Dim found As Boolean
        found = False
        Dim slot As Long, slotRow As Long
        For slot = 0 To CAMPAIGNS_NUM_SLOTS - 1
            slotRow = CAMPAIGNS_FIRST_MAIN_ROW + slot * CAMPAIGNS_SLOT_SIZE
            If Trim$(campaignsWs.Cells(slotRow, CAMPAIGNS_COL_NAME).Value & "") = typedName Then
                ws.Cells(r, RT_COL_LINKED_CAMPAIGN_KEY).Value = campaignsWs.Cells(slotRow, CAMPAIGNS_COL_KEY).Value
                found = True
                Exit For
            End If
        Next slot
        If Not found Then
            ws.Cells(r, RT_COL_LINKED_CAMPAIGN_KEY).Value = ""
        End If
    End If

    Application.EnableEvents = prevEvents
    HandleLinkedCampaignChange = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandleLinkedCampaignChange = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Sub RefreshLinkedCampaignLabels(ws As Worksheet)
    Dim campaignsWs As Worksheet, campaignsArchiveWs As Worksheet
    Set campaignsWs = ThisWorkbook.Sheets("Campaigns")
    Set campaignsArchiveWs = ThisWorkbook.Sheets("Campaigns Archive")

    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, RT_COL_KEY, RT_FIRST_DATA_ROW)
    If lastRow < RT_FIRST_DATA_ROW Then Exit Sub

    Dim r As Long
    For r = RT_FIRST_DATA_ROW To lastRow
        Dim keyText As String
        keyText = Trim$(ws.Cells(r, RT_COL_LINKED_CAMPAIGN_KEY).Value & "")
        If Len(keyText) > 0 Then
            Dim keyVal As Double
            keyVal = CDbl(keyText)

            Dim campaignName As String
            campaignName = ""

            Dim slot As Long, slotRow As Long
            For slot = 0 To CAMPAIGNS_NUM_SLOTS - 1
                slotRow = CAMPAIGNS_FIRST_MAIN_ROW + slot * CAMPAIGNS_SLOT_SIZE
                If Len(Trim$(campaignsWs.Cells(slotRow, CAMPAIGNS_COL_KEY).Value & "")) > 0 Then
                    If CDbl(campaignsWs.Cells(slotRow, CAMPAIGNS_COL_KEY).Value) = keyVal Then
                        campaignName = campaignsWs.Cells(slotRow, CAMPAIGNS_COL_NAME).Value
                        Exit For
                    End If
                End If
            Next slot

            If Len(campaignName) = 0 Then
                Dim archLastRow As Long
                archLastRow = LastHierarchyRow(campaignsArchiveWs, CAMPAIGNS_ARCHIVE_COL_KEY, 2)
                Dim ar As Long
                For ar = 2 To archLastRow
                    If Len(Trim$(campaignsArchiveWs.Cells(ar, CAMPAIGNS_ARCHIVE_COL_KEY).Value & "")) > 0 Then
                        If CDbl(campaignsArchiveWs.Cells(ar, CAMPAIGNS_ARCHIVE_COL_KEY).Value) = keyVal Then
                            campaignName = campaignsArchiveWs.Cells(ar, CAMPAIGNS_ARCHIVE_COL_NAME).Value
                            Exit For
                        End If
                    End If
                Next ar
            End If

            If Len(campaignName) > 0 Then
                ws.Cells(r, RT_COL_LINKED_CAMPAIGN).Value = campaignName
            End If
        End If
    Next r
End Sub

Public Function HandleTaskCollapseDoubleClick(r As Long) As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = RTSheet()

    Dim numSubtasks As Long
    numSubtasks = 0
    If Len(Trim$(ws.Cells(r, RT_COL_NUM_SUBTASKS).Value & "")) > 0 Then
        numSubtasks = CLng(ws.Cells(r, RT_COL_NUM_SUBTASKS).Value)
    End If

    If numSubtasks = 0 Then
        HandleTaskCollapseDoubleClick = "NO_SUBTASKS"
        Exit Function
    End If

    Dim myLevel As Long
    myLevel = CLng(ws.Cells(r, RT_COL_LEVEL).Value)

    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, RT_COL_KEY, RT_FIRST_DATA_ROW)

    Dim blockEnd As Long
    blockEnd = r
    Dim i As Long
    For i = r + 1 To lastRow
        If CLng(ws.Cells(i, RT_COL_LEVEL).Value) > myLevel Then
            blockEnd = i
        Else
            Exit For
        End If
    Next i

    If blockEnd > r Then
        Dim isCurrentlyHidden As Boolean
        isCurrentlyHidden = ws.Rows(r + 1).Hidden
        ws.Rows((r + 1) & ":" & blockEnd).Hidden = Not isCurrentlyHidden
        HandleTaskCollapseDoubleClick = IIf(isCurrentlyHidden, "EXPANDED", "COLLAPSED")
    Else
        HandleTaskCollapseDoubleClick = "NO_SUBTASKS"
    End If
    Exit Function
ErrHandler:
    HandleTaskCollapseDoubleClick = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function HandleSiblingReorder(r As Long, direction As String) As String
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = RTSheet()

    Dim myParentKey As String
    myParentKey = Trim$(ws.Cells(r, RT_COL_PARENT_KEY).Value & "")
    Dim myLevel As Long
    myLevel = CLng(ws.Cells(r, RT_COL_LEVEL).Value)

    Dim targetRow As Long
    targetRow = 0

    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, RT_COL_KEY, RT_FIRST_DATA_ROW)

    If LCase$(direction) = "up" Then
        Dim i As Long
        For i = r - 1 To RT_FIRST_DATA_ROW Step -1
            If CLng(ws.Cells(i, RT_COL_LEVEL).Value) < myLevel Then Exit For
            If CLng(ws.Cells(i, RT_COL_LEVEL).Value) = myLevel And Trim$(ws.Cells(i, RT_COL_PARENT_KEY).Value & "") = myParentKey Then
                targetRow = i
                Exit For
            End If
        Next i
    Else
        Dim j As Long
        For j = r + 1 To lastRow
            If CLng(ws.Cells(j, RT_COL_LEVEL).Value) < myLevel Then Exit For
            If CLng(ws.Cells(j, RT_COL_LEVEL).Value) = myLevel And Trim$(ws.Cells(j, RT_COL_PARENT_KEY).Value & "") = myParentKey Then
                targetRow = j
                Exit For
            End If
        Next j
    End If

    If targetRow = 0 Then
        Application.EnableEvents = prevEvents
        HandleSiblingReorder = "NOOP"
        Exit Function
    End If

    Dim myOrder As Double, targetOrder As Double
    myOrder = CDbl(ws.Cells(r, RT_COL_SIBLING_ORDER).Value)
    targetOrder = CDbl(ws.Cells(targetRow, RT_COL_SIBLING_ORDER).Value)

    ws.Cells(r, RT_COL_SIBLING_ORDER).Value = targetOrder
    ws.Cells(targetRow, RT_COL_SIBLING_ORDER).Value = myOrder

    RenumberHierarchy ws, RTColMap()
    RefreshButtonStyling ws
    RefreshLinkedCampaignLabels ws

    Application.EnableEvents = prevEvents
    HandleSiblingReorder = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandleSiblingReorder = "ERROR " & Err.Number & ": " & Err.Description
End Function

Private Function FindRowByID(ws As Worksheet, idText As String) As Long
    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, RT_COL_KEY, RT_FIRST_DATA_ROW)
    Dim r As Long
    For r = RT_FIRST_DATA_ROW To lastRow
        If Trim$(ws.Cells(r, RT_COL_ID).Value & "") = idText Then
            FindRowByID = r
            Exit Function
        End If
    Next r
    FindRowByID = 0
End Function

Private Function FindRowByKey(ws As Worksheet, recurringKey As Double) As Long
    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, RT_COL_KEY, RT_FIRST_DATA_ROW)
    Dim r As Long
    For r = RT_FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, RT_COL_KEY).Value & "")) > 0 Then
            If CDbl(ws.Cells(r, RT_COL_KEY).Value) = recurringKey Then
                FindRowByKey = r
                Exit Function
            End If
        End If
    Next r
    FindRowByKey = 0
End Function

' Public for the same reason as modTasks.WouldCreateCycle -- direct test
' access, bypassing HandleSubtaskOfChange's MsgBox.
Public Function WouldCreateCycle(ws As Worksheet, movingKey As Double, proposedParentKey As String) As Boolean
    Dim currentKey As String
    currentKey = proposedParentKey
    Dim guard As Long
    guard = 0
    Do While Len(currentKey) > 0 And guard < 10000
        If CDbl(currentKey) = movingKey Then
            WouldCreateCycle = True
            Exit Function
        End If
        Dim parentRow As Long
        parentRow = FindRowByKey(ws, CDbl(currentKey))
        If parentRow = 0 Then Exit Do
        currentKey = Trim$(ws.Cells(parentRow, RT_COL_PARENT_KEY).Value & "")
        guard = guard + 1
    Loop
    WouldCreateCycle = False
End Function

' ============================================================================
' RegenerateRecurringTasks (§12.1)
' ============================================================================

' Computes every scheduled occurrence date in (fromDate, throughDate]
' (fromDate itself excluded -- it's the watermark, already generated) for
' the given Recurrence/Task Start On. Public and pure (no sheet writes, no
' MsgBox) specifically so the date math is independently testable.
' outDates is 1-based, sized to outCount; a generous 10000-slot working
' buffer comfortably covers any realistic gap (~27 years of Daily) without
' needing a resizable-array dance mid-loop.
Public Sub ComputeScheduledDates(recurrence As String, taskStartOn As String, fromDate As Date, throughDate As Date, ByRef outDates() As Date, ByRef outCount As Long)
    outCount = 0
    ReDim outDates(1 To 10000)

    If fromDate > throughDate Then Exit Sub

    Dim d As Date
    Select Case recurrence
        Case "Daily"
            ' Weekdays only (user-reported/confirmed): a work task tracker
            ' has no use for Saturday/Sunday occurrences, which the user
            ' had to delete by hand every week -- "Daily" now means
            ' Mon-Fri, not every calendar day.
            For d = fromDate To throughDate
                If Weekday(d, vbMonday) <= 5 Then
                    outCount = outCount + 1
                    outDates(outCount) = d
                End If
            Next d
        Case "Weekly"
            Dim targetWeekday As Long
            targetWeekday = WeekdayNumberFromName(taskStartOn)
            If targetWeekday > 0 Then
                For d = fromDate To throughDate
                    If Weekday(d, vbMonday) = targetWeekday Then
                        outCount = outCount + 1
                        outDates(outCount) = d
                    End If
                Next d
            End If
        Case "Monthly"
            If Len(Trim$(taskStartOn)) > 0 And IsNumeric(taskStartOn) Then
                Dim targetDay As Long
                targetDay = CLng(taskStartOn)
                For d = fromDate To throughDate
                    Dim lastDayOfMonth As Long
                    lastDayOfMonth = Day(DateSerial(Year(d), Month(d) + 1, 0))
                    Dim effectiveDay As Long
                    effectiveDay = targetDay
                    If effectiveDay > lastDayOfMonth Then effectiveDay = lastDayOfMonth
                    If Day(d) = effectiveDay Then
                        outCount = outCount + 1
                        outDates(outCount) = d
                    End If
                Next d
            End If
    End Select
End Sub

Private Function WeekdayNumberFromName(name As String) As Long
    Select Case name
        Case "Monday": WeekdayNumberFromName = 1
        Case "Tuesday": WeekdayNumberFromName = 2
        Case "Wednesday": WeekdayNumberFromName = 3
        Case "Thursday": WeekdayNumberFromName = 4
        Case "Friday": WeekdayNumberFromName = 5
        Case "Saturday": WeekdayNumberFromName = 6
        Case "Sunday": WeekdayNumberFromName = 7
        Case Else: WeekdayNumberFromName = 0
    End Select
End Function

' Generates (and writes into Tasks) every due occurrence for one root
' template, advances its watermark to TODAY() regardless of how many
' occurrences actually generated (so a template with zero matches, or one
' capped by MAX_OCCURRENCES_PER_RUN, is never retried for the same gap),
' and returns how many older occurrences were skipped by the safety cap
' (0 if none). Public and MsgBox-free specifically so this -- the actual
' risky logic -- is directly testable; RegenerateRecurringTasks (below)
' is the thin, MsgBox-showing production wrapper around it, same
' reasoning as CascadeDelete/WouldCreateCycle's split from
' HandleStatusChange/HandleSubtaskOfChange.
Public Function ProcessTemplate(ws As Worksheet, tasksWs As Worksheet, templateRootRow As Long) As Long
    Dim recurrence As String, taskStartOn As String
    recurrence = Trim$(ws.Cells(templateRootRow, RT_COL_RECURRENCE).Value & "")
    taskStartOn = Trim$(ws.Cells(templateRootRow, RT_COL_TASK_START_ON).Value & "")

    Dim watermarkRaw As String
    watermarkRaw = Trim$(ws.Cells(templateRootRow, RT_COL_LAST_GENERATED).Value & "")

    Dim fromDate As Date
    If Len(watermarkRaw) = 0 Then
        ' Defensive fallback -- HandleTaskNameChange always seeds this on
        ' creation, so a genuinely blank watermark shouldn't occur in
        ' practice. Treat as "nothing to backfill" rather than guessing.
        fromDate = Date + 1
    Else
        fromDate = CDate(ws.Cells(templateRootRow, RT_COL_LAST_GENERATED).Value) + 1
    End If

    Dim allDates() As Date, dateCount As Long
    ComputeScheduledDates recurrence, taskStartOn, fromDate, Date, allDates, dateCount

    ' "If Missed = Skip" (V4, PLAN.md §22.4): first archive any of this
    ' template's still-open instances whose deadline day has passed, then
    ' only generate occurrences whose deadline day is today or later -- an
    ' instance that would be skipped the moment it appeared is never
    ' created at all.
    If Trim$(ws.Cells(templateRootRow, RT_COL_IF_MISSED).Value & "") = "Skip" Then
        modRecurringTasksWriter.SkipMissedInstances CDbl(ws.Cells(templateRootRow, RT_COL_KEY).Value)

        Dim rootOffset As Double
        rootOffset = Val(ws.Cells(templateRootRow, RT_COL_DEADLINE_OFFSET).Value & "")
        Dim kept As Long, q As Long
        kept = 0
        For q = 1 To dateCount
            If allDates(q) + rootOffset >= Date Then
                kept = kept + 1
                allDates(kept) = allDates(q)
            End If
        Next q
        dateCount = kept
    End If

    Dim skipped As Long
    Dim genCount As Long
    skipped = 0
    If dateCount > MAX_OCCURRENCES_PER_RUN Then
        skipped = dateCount - MAX_OCCURRENCES_PER_RUN
        genCount = MAX_OCCURRENCES_PER_RUN
    Else
        genCount = dateCount
    End If

    Dim startIdx As Long
    startIdx = dateCount - genCount + 1   ' the most recent genCount dates

    Dim i As Long
    For i = startIdx To dateCount
        GenerateInstance ws, tasksWs, templateRootRow, allDates(i)
    Next i

    If genCount > 0 Then
        ' modRecurringTasksWriter.FinalizeGeneration does the variant's own
        ' RenumberHierarchy + RefreshButtonStyling -- GenerateInstance()
        ' writes Tasks' Up/Down (and, employee only, Timer) cells directly,
        ' bypassing HandleTaskNameChange entirely, so the button-look those
        ' generated rows need has to be applied here explicitly too, not
        ' just at modTasks.bas's/modTasksManager.bas's own RenumberHierarchy
        ' call sites.
        modRecurringTasksWriter.FinalizeGeneration tasksWs
    End If

    ws.Cells(templateRootRow, RT_COL_LAST_GENERATED).Value = Date

    ProcessTemplate = skipped
End Function

' Deep-copies one occurrence of templateRootRow's branch into Tasks, dated
' occurrenceDate. Relies on Recurring Tasks' own tree already being in
' physical depth-first order (guaranteed by RenumberHierarchy's own
' invariant -- see modHierarchy.bas), so a plain top-to-bottom scan of the
' branch visits each parent before its children.
'
' V4 (PLAN.md §22.3) -- three changes from the original "copy the whole
' branch verbatim" behavior:
'   1. Per-row schedules: a subtask is only copied if its own Status is
'      Active, its parent was copied, and its own Recurrence is blank
'      ("same as parent") or matches occurrenceDate. The root's schedule
'      gates everything below it.
'   2. Time: Deadline's time of day comes from the row's Time column
'      (else the old fixed 18:00), and each copied parent's children are
'      ordered by Time (untimed last, ties keep template order).
'   3. Placement (user-reported): the generated root used to inherit the
'      TEMPLATE's own SiblingOrder, colliding with the user's existing
'      top-level tasks (generated "1" sorted beside/above real task "1").
'      It now takes the next free root SiblingOrder in Tasks, so a new
'      instance always lands below everything already there.
Private Sub GenerateInstance(ws As Worksheet, tasksWs As Worksheet, templateRootRow As Long, occurrenceDate As Date)
    Dim idText As String
    idText = Trim$(ws.Cells(templateRootRow, RT_COL_ID).Value & "")

    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, RT_COL_KEY, RT_FIRST_DATA_ROW)

    Dim branchRows() As Long
    Dim n As Long
    n = 0
    ReDim branchRows(1 To 500)
    n = n + 1
    branchRows(n) = templateRootRow

    Dim i As Long
    For i = RT_FIRST_DATA_ROW To lastRow
        If i <> templateRootRow Then
            Dim otherID As String
            otherID = Trim$(ws.Cells(i, RT_COL_ID).Value & "")
            If Left$(otherID, Len(idText) + 1) = idText & "." Then
                n = n + 1
                branchRows(n) = i
            End If
        End If
    Next i

    ' --- Pass 1: which rows belong to this occurrence ---
    Dim included() As Boolean
    ReDim included(1 To n)
    Dim includedByKey As Object
    Set includedByKey = CreateObject("Scripting.Dictionary")

    Dim k As Long
    For k = 1 To n
        Dim srcRow As Long
        srcRow = branchRows(k)
        If k = 1 Then
            included(k) = True
        Else
            Dim parentKeyText As String
            parentKeyText = Trim$(ws.Cells(srcRow, RT_COL_PARENT_KEY).Value & "")
            included(k) = False
            If includedByKey.Exists(parentKeyText) Then
                If includedByKey(parentKeyText) Then
                    If Trim$(ws.Cells(srcRow, RT_COL_STATUS).Value & "") = "Active" Then
                        included(k) = RowScheduleMatches(ws, srcRow, occurrenceDate)
                    End If
                End If
            End If
        End If
        includedByKey(Trim$(ws.Cells(srcRow, RT_COL_KEY).Value & "")) = included(k)
    Next k

    ' --- Pass 2: time of day + per-parent ordering by time ---
    Dim timeOfDay() As Double
    Dim sortKey() As Double
    ReDim timeOfDay(1 To n)
    ReDim sortKey(1 To n)
    For k = 1 To n
        timeOfDay(k) = TimeOfDayOrNegative(ws.Cells(branchRows(k), RT_COL_TIME).Value)
        Dim minutesKey As Double
        If timeOfDay(k) >= 0 Then
            minutesKey = Round(timeOfDay(k) * 1440, 0)
        Else
            minutesKey = 2000   ' untimed sorts after every real time (max 1439)
        End If
        sortKey(k) = minutesKey * 100000# + Val(ws.Cells(branchRows(k), RT_COL_SIBLING_ORDER).Value & "")
    Next k

    Dim newSiblingOrder() As Double
    ReDim newSiblingOrder(1 To n)
    Dim j As Long
    For k = 2 To n
        If included(k) Then
            Dim myParent As String
            myParent = Trim$(ws.Cells(branchRows(k), RT_COL_PARENT_KEY).Value & "")
            Dim rank As Long
            rank = 1
            For j = 2 To n
                If j <> k And included(j) Then
                    If Trim$(ws.Cells(branchRows(j), RT_COL_PARENT_KEY).Value & "") = myParent Then
                        If sortKey(j) < sortKey(k) Or (sortKey(j) = sortKey(k) And j < k) Then rank = rank + 1
                    End If
                End If
            Next j
            newSiblingOrder(k) = rank * 10
        End If
    Next k
    newSiblingOrder(1) = NextSiblingOrder(tasksWs, TasksKeyColMap(), "")

    ' --- Pass 3: write the included rows into Tasks ---
    Dim keyMap As Object
    Set keyMap = CreateObject("Scripting.Dictionary")

    Dim tasksLastRow As Long
    tasksLastRow = LastHierarchyRow(tasksWs, TASKS_COL_TASK_KEY, TASKS_FIRST_DATA_ROW)
    Dim destRow As Long
    If tasksLastRow < TASKS_FIRST_DATA_ROW Then
        destRow = TASKS_FIRST_DATA_ROW
    Else
        destRow = tasksLastRow + 1
    End If

    For k = 1 To n
        If included(k) Then
            srcRow = branchRows(k)

            Dim newTaskKey As Double
            newTaskKey = modRecurringTasksWriter.NextGeneratedTaskKey()
            keyMap(Trim$(ws.Cells(srcRow, RT_COL_KEY).Value & "")) = newTaskKey

            Dim newParentKeyText As String
            Dim rawParent As String
            rawParent = Trim$(ws.Cells(srcRow, RT_COL_PARENT_KEY).Value & "")
            If k = 1 Or Len(rawParent) = 0 Then
                newParentKeyText = ""
            Else
                newParentKeyText = CStr(keyMap(rawParent))
            End If

            tasksWs.Cells(destRow, TASKS_COL_TASK).Value = ws.Cells(srcRow, RT_COL_TASK).Value

            ' "None", not whatever the template's own Linked Campaign happens
            ' to hold, if that's blank (defensive -- HandleTaskNameChange
            ' already defaults a template's own Linked Campaign to "None" on
            ' creation, so this should rarely actually trigger).
            Dim linkedCampaignVal As String
            linkedCampaignVal = Trim$(ws.Cells(srcRow, RT_COL_LINKED_CAMPAIGN).Value & "")
            If Len(linkedCampaignVal) = 0 Then linkedCampaignVal = "None"

            ' Now(), not occurrenceDate (user-reported, PLAN.md §19 Phase 5):
            ' Start Date's meaning everywhere else in this workbook is "when
            ' was this row created" -- occurrenceDate is still what Deadline
            ' below is based on, so a backfilled batch keeps distinct days.
            tasksWs.Cells(destRow, TASKS_COL_START_DATE).Value = Now()

            Dim offsetDays As Double
            offsetDays = 0
            If Len(Trim$(ws.Cells(srcRow, RT_COL_DEADLINE_OFFSET).Value & "")) > 0 Then
                offsetDays = CDbl(ws.Cells(srcRow, RT_COL_DEADLINE_OFFSET).Value)
            End If
            ' The row's own Time if set (§22.3), else 18:00 as the default
            ' end-of-workday cutoff (PLAN.md §19 Phase 5).
            Dim deadlineVal As Date
            If timeOfDay(k) >= 0 Then
                deadlineVal = occurrenceDate + offsetDays + timeOfDay(k)
            Else
                deadlineVal = occurrenceDate + offsetDays + TimeSerial(18, 0, 0)
            End If

            ' Project/Linked Campaign/Linked Campaign Key/Priority/Deadline/
            ' Status genuinely differ in column position between variants (and
            ' Timer/Time Spent don't exist for manager) -- delegated to
            ' modRecurringTasksWriter, see this module's own column-layout
            ' comment above for why.
            modRecurringTasksWriter.WriteVariantSpecificFields tasksWs, destRow, _
                ws.Cells(srcRow, RT_COL_PROJECT).Value, linkedCampaignVal, _
                ws.Cells(srcRow, RT_COL_LINKED_CAMPAIGN_KEY).Value, ws.Cells(srcRow, RT_COL_PRIORITY).Value, deadlineVal

            tasksWs.Cells(destRow, TASKS_COL_NOTES).Value = ws.Cells(srcRow, RT_COL_NOTES).Value

            ' Up/Down defaults (user-reported, PLAN.md §19 Phase 5): this
            ' writes a Tasks row directly, bypassing HandleTaskNameChange, so
            ' its "new task" defaults are applied here explicitly. Timer's
            ' own default (employee only) is written by
            ' WriteVariantSpecificFields above instead.
            tasksWs.Cells(destRow, TASKS_COL_UP).Value = ICON_UP
            tasksWs.Cells(destRow, TASKS_COL_DOWN).Value = ICON_DOWN

            tasksWs.Cells(destRow, TASKS_COL_TASK_KEY).Value = newTaskKey
            tasksWs.Cells(destRow, TASKS_COL_PARENT_KEY).Value = newParentKeyText
            tasksWs.Cells(destRow, TASKS_COL_LEVEL).Value = ws.Cells(srcRow, RT_COL_LEVEL).Value
            tasksWs.Cells(destRow, TASKS_COL_SIBLING_ORDER).Value = newSiblingOrder(k)
            If k = 1 Then
                modRecurringTasksWriter.WriteRecurringSource tasksWs, destRow, CDbl(ws.Cells(srcRow, RT_COL_KEY).Value)
            End If

            destRow = destRow + 1
        End If
    Next k
End Sub

' Does row r's OWN Recurrence/Task Start On match date d? Blank
' Recurrence means "same as parent" (PLAN.md §22.3) -- the caller has
' already established the parent matched, so blank is a match.
Private Function RowScheduleMatches(ws As Worksheet, r As Long, d As Date) As Boolean
    Dim rec As String
    rec = Trim$(ws.Cells(r, RT_COL_RECURRENCE).Value & "")
    If Len(rec) = 0 Then
        RowScheduleMatches = True
        Exit Function
    End If
    Dim dates() As Date, cnt As Long
    ComputeScheduledDates rec, Trim$(ws.Cells(r, RT_COL_TASK_START_ON).Value & ""), d, d, dates, cnt
    RowScheduleMatches = (cnt > 0)
End Function

' A Time cell's time of day as a 0..1 day fraction, or -1 if blank/not a
' time. Accepts a real time value (the normal case -- Excel parses a
' typed "10:00") or a text "10:00" (a cell formatted as Text).
Private Function TimeOfDayOrNegative(v As Variant) As Double
    TimeOfDayOrNegative = -1
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(v & "")) = 0 Then Exit Function
    On Error GoTo NotATime
    Dim x As Double
    If IsNumeric(v) Then
        x = CDbl(v)
    Else
        x = CDbl(CDate(v))
    End If
    TimeOfDayOrNegative = x - Int(x)
    Exit Function
NotATime:
    TimeOfDayOrNegative = -1
End Function

' The slice of Tasks' column map NextSiblingOrder needs (key, parent key,
' sibling order) -- those three sit at the same columns in both variants,
' see the TASKS_COL_* comment above.
Private Function TasksKeyColMap() As HierarchyColumnMap
    Dim m As HierarchyColumnMap
    m.FirstDataRow = TASKS_FIRST_DATA_ROW
    m.ColKey = TASKS_COL_TASK_KEY
    m.ColParentKey = TASKS_COL_PARENT_KEY
    m.ColSiblingOrder = TASKS_COL_SIBLING_ORDER
    TasksKeyColMap = m
End Function

' Production entry point -- called from Workbook_Open (a later Phase 3
' item). Walks every Active, root-level (RecurringParentKey blank)
' template and processes it; warns once per template that had occurrences
' skipped by the safety cap. Swallows its own errors (never lets an
' unhandled runtime error surface a blocking modal on file open, per
' CLAUDE.md's rule for anything invoked this early/automatically).
Public Sub RegenerateRecurringTasks()
    On Error GoTo ErrHandler
    Dim ws As Worksheet, tasksWs As Worksheet
    Set ws = RTSheet()
    Set tasksWs = ThisWorkbook.Sheets("Tasks")

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, RT_COL_KEY, RT_FIRST_DATA_ROW)

    Dim r As Long
    For r = RT_FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, RT_COL_KEY).Value & "")) > 0 Then
            If Len(Trim$(ws.Cells(r, RT_COL_PARENT_KEY).Value & "")) = 0 Then
                If Trim$(ws.Cells(r, RT_COL_STATUS).Value & "") = "Active" Then
                    Dim taskName As String
                    taskName = ws.Cells(r, RT_COL_TASK).Value

                    Dim skipped As Long
                    skipped = ProcessTemplate(ws, tasksWs, r)

                    If skipped > 0 Then
                        MsgBox "'" & taskName & "' had " & skipped & " older occurrence(s) beyond the " & _
                               MAX_OCCURRENCES_PER_RUN & "-occurrence catch-up limit -- those were skipped.", _
                               vbInformation, "Recurring Task Backfill Limited"
                    End If
                End If
            End If
        End If
    Next r

    Application.EnableEvents = prevEvents
    Exit Sub
ErrHandler:
    Application.EnableEvents = prevEvents
End Sub
