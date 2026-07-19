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
Private Const RT_COL_DEADLINE_OFFSET As Long = 13
Private Const RT_COL_NOTES As Long = 14
Private Const RT_COL_KEY As Long = 15             ' RecurringKey
Private Const RT_COL_PARENT_KEY As Long = 16       ' RecurringParentKey
Private Const RT_COL_LEVEL As Long = 17
Private Const RT_COL_SIBLING_ORDER As Long = 18
Private Const RT_COL_LINKED_CAMPAIGN_KEY As Long = 19
Private Const RT_COL_LAST_GENERATED As Long = 20   ' LastGeneratedThroughDate, root rows only
Private Const RT_LAST_COL As Long = 20

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
' isn't shared directly) ---
Private Const TASKS_COL_UP As Long = 2
Private Const TASKS_COL_DOWN As Long = 3
Private Const TASKS_COL_START_DATE As Long = 4
Private Const TASKS_COL_TASK As Long = 5
Private Const TASKS_COL_PROJECT As Long = 8
Private Const TASKS_COL_LINKED_CAMPAIGN As Long = 9
Private Const TASKS_COL_PRIORITY As Long = 10
Private Const TASKS_COL_DEADLINE As Long = 11
Private Const TASKS_COL_STATUS As Long = 12
Private Const TASKS_COL_TIMER As Long = 14
Private Const TASKS_COL_TIME_SPENT As Long = 15
Private Const TASKS_COL_NOTES As Long = 16
Private Const TASKS_COL_TASK_KEY As Long = 17
Private Const TASKS_COL_PARENT_KEY As Long = 18
Private Const TASKS_COL_LEVEL As Long = 19
Private Const TASKS_COL_SIBLING_ORDER As Long = 20
Private Const TASKS_COL_LINKED_CAMPAIGN_KEY As Long = 22
Private Const TASKS_FIRST_DATA_ROW As Long = 2

Private Const MAX_OCCURRENCES_PER_RUN As Long = 30   ' tunable, PLAN.md §5/§12.1

Private Function RTColMap() As HierarchyColumnMap
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
            For d = fromDate To throughDate
                outCount = outCount + 1
                outDates(outCount) = d
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
        RenumberHierarchy tasksWs, modTasks.TasksColMap()
        ' GenerateInstance() writes Tasks' Up/Down/Timer cells directly,
        ' bypassing HandleTaskNameChange entirely (see its own comment
        ' below) -- so the button-look those generated rows need has to
        ' be applied here explicitly too, not just at modTasks.bas's own
        ' RenumberHierarchy call sites.
        modTasks.RefreshButtonStyling tasksWs
    End If

    ws.Cells(templateRootRow, RT_COL_LAST_GENERATED).Value = Date

    ProcessTemplate = skipped
End Function

' Deep-copies templateRootRow's entire branch (root + all descendants) as
' one new instance in Tasks, dated occurrenceDate. Relies on Recurring
' Tasks' own tree already being in physical depth-first order (guaranteed
' by RenumberHierarchy's own invariant -- see modHierarchy.bas), so a
' plain top-to-bottom scan of the branch visits each parent before its
' children, letting the RecurringKey -> new-TaskKey map be built in a
' single forward pass.
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

    Dim k As Long
    For k = 1 To n
        Dim srcRow As Long
        srcRow = branchRows(k)

        Dim newTaskKey As Double
        newTaskKey = modTasks.NextTaskKey()
        keyMap(CDbl(ws.Cells(srcRow, RT_COL_KEY).Value)) = newTaskKey

        Dim newParentKeyText As String
        Dim rawParent As String
        rawParent = Trim$(ws.Cells(srcRow, RT_COL_PARENT_KEY).Value & "")
        If Len(rawParent) = 0 Then
            newParentKeyText = ""
        Else
            newParentKeyText = CStr(keyMap(CDbl(rawParent)))
        End If

        tasksWs.Cells(destRow, TASKS_COL_TASK).Value = ws.Cells(srcRow, RT_COL_TASK).Value
        tasksWs.Cells(destRow, TASKS_COL_PROJECT).Value = ws.Cells(srcRow, RT_COL_PROJECT).Value

        ' "None", not whatever the template's own Linked Campaign happens
        ' to hold, if that's blank (defensive -- HandleTaskNameChange
        ' already defaults a template's own Linked Campaign to "None" on
        ' creation, so this should rarely actually trigger).
        Dim linkedCampaignVal As String
        linkedCampaignVal = Trim$(ws.Cells(srcRow, RT_COL_LINKED_CAMPAIGN).Value & "")
        If Len(linkedCampaignVal) = 0 Then linkedCampaignVal = "None"
        tasksWs.Cells(destRow, TASKS_COL_LINKED_CAMPAIGN).Value = linkedCampaignVal
        tasksWs.Cells(destRow, TASKS_COL_LINKED_CAMPAIGN_KEY).Value = ws.Cells(srcRow, RT_COL_LINKED_CAMPAIGN_KEY).Value
        tasksWs.Cells(destRow, TASKS_COL_PRIORITY).Value = ws.Cells(srcRow, RT_COL_PRIORITY).Value
        ' Now(), not occurrenceDate (user-reported, PLAN.md §19 Phase 5):
        ' occurrenceDate has no time component (a plain Date from
        ' ComputeScheduledDates' loop), so Start Date always showed
        ' midnight instead of the actual moment this instance was
        ' generated -- inconsistent with every manually-created task,
        ' whose Start Date is Now() (modTasks.bas's HandleTaskNameChange).
        ' occurrenceDate is intentionally still used for Deadline below --
        ' that field needs to represent which CALENDAR DAY this instance
        ' is actually for (so a backfilled batch's occurrences keep
        ' distinct deadlines), whereas Start Date's meaning everywhere
        ' else in this workbook is "when was this row created," not
        ' "which day does it conceptually represent."
        tasksWs.Cells(destRow, TASKS_COL_START_DATE).Value = Now()

        Dim offsetDays As Double
        offsetDays = 0
        If Len(Trim$(ws.Cells(srcRow, RT_COL_DEADLINE_OFFSET).Value & "")) > 0 Then
            offsetDays = CDbl(ws.Cells(srcRow, RT_COL_DEADLINE_OFFSET).Value)
        End If
        ' + TimeSerial(18,0,0), not just the bare date (user-requested,
        ' PLAN.md §19 Phase 5): occurrenceDate has no time component, so
        ' the Deadline always landed at midnight -- an unhelpfully early
        ' cutoff for a task whose actual working day hasn't started yet.
        ' 18:00 as a reasonable default end-of-workday time.
        tasksWs.Cells(destRow, TASKS_COL_DEADLINE).Value = occurrenceDate + offsetDays + TimeSerial(18, 0, 0)

        tasksWs.Cells(destRow, TASKS_COL_STATUS).Value = "Not Started"
        tasksWs.Cells(destRow, TASKS_COL_NOTES).Value = ws.Cells(srcRow, RT_COL_NOTES).Value

        ' Up/Down/Timer/Time Spent defaults (user-reported, PLAN.md §19
        ' Phase 5): GenerateInstance writes a Tasks row directly, bypassing
        ' modTasks.HandleTaskNameChange entirely, so it never picked up
        ' any of the "new task" defaults added there over the course of
        ' this phase -- a generated instance looked all-blank/undefaulted
        ' until manually double-clicked, unlike a normally-typed task.
        tasksWs.Cells(destRow, TASKS_COL_UP).Value = ICON_UP
        tasksWs.Cells(destRow, TASKS_COL_DOWN).Value = ICON_DOWN
        tasksWs.Cells(destRow, TASKS_COL_TIMER).Value = ICON_TIMER_RUNNING & " Start"
        ' 0, not blank -- Time Spent (hrs) is hidden, but its value drives
        ' the visible friendly "Time Spent" column's formula
        ' (=IF(ts="","",...)), so a genuinely blank hrs cell showed as a
        ' blank friendly cell instead of "0h 00m".
        tasksWs.Cells(destRow, TASKS_COL_TIME_SPENT).Value = 0

        tasksWs.Cells(destRow, TASKS_COL_TASK_KEY).Value = newTaskKey
        tasksWs.Cells(destRow, TASKS_COL_PARENT_KEY).Value = newParentKeyText
        tasksWs.Cells(destRow, TASKS_COL_LEVEL).Value = ws.Cells(srcRow, RT_COL_LEVEL).Value
        tasksWs.Cells(destRow, TASKS_COL_SIBLING_ORDER).Value = ws.Cells(srcRow, RT_COL_SIBLING_ORDER).Value

        destRow = destRow + 1
    Next k
End Sub

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
