Attribute VB_Name = "modTasks"
Option Explicit

' Tasks sheet business logic (PLAN.md §6.3). Every Public Function here is
' independently callable via Application.Run for testing (per CLAUDE.md's
' rule: every such entry point has its own On Error GoTo ErrHandler and
' never lets an error propagate unhandled -- an unhandled VBA runtime
' error pops a blocking modal dialog regardless of Application.Visible/
' DisplayAlerts). The actual Worksheet_Change/Worksheet_BeforeDoubleClick
' event Subs are thin dispatchers living in vba/Tasks.bas (injected
' directly into the Tasks worksheet's own code module by Phase 4, not a
' standard module) -- they just route by Target.Column/Row into the
' Functions here.
'
' Reuses modHierarchy's generic RenumberHierarchy/NextSiblingOrder (Phase
' 1) and frmNotes (Phase 1, called directly from Tasks.bas, not from
' here, since ShowForCell's modal Me.Show must never run during
' automated testing -- see frmNotes' own docstring).

' --- Tasks column layout (sheets/tasks.py) ---------------------------------
Private Const COL_ID As Long = 1
Private Const COL_UP As Long = 2
Private Const COL_DOWN As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_TASK As Long = 5
Private Const COL_NUM_SUBTASKS As Long = 6
Private Const COL_SUBTASK_OF As Long = 7
Private Const COL_PROJECT As Long = 8
Private Const COL_LINKED_CAMPAIGN As Long = 9
Private Const COL_PRIORITY As Long = 10
Private Const COL_DEADLINE As Long = 11
Private Const COL_STATUS As Long = 12
Private Const COL_DAYS_LEFT As Long = 13
Private Const COL_TIMER As Long = 14
Private Const COL_TIME_SPENT As Long = 15
Private Const COL_NOTES As Long = 16
Private Const COL_TASK_KEY As Long = 17
Private Const COL_PARENT_KEY As Long = 18
Private Const COL_LEVEL As Long = 19
Private Const COL_SIBLING_ORDER As Long = 20
Private Const COL_TIMER_RUNNING_SINCE As Long = 21
Private Const COL_LINKED_CAMPAIGN_KEY As Long = 22
' Columns 23-26 (SortKey/CampaignSortHelper/CampaignRank/CampaignRankKey)
' are deliberately excluded from TasksColMap's LastCol below -- they're
' per-row formulas referencing that same row's own cells plus fixed
' absolute ranges, so they self-heal as data moves underneath them during
' RenumberHierarchy's clear-and-rewrite; including them would instead
' freeze their currently-cached values as static numbers, destroying them.
' Named here (previously only used as bare literals in sheets/tasks.py)
' because the row extender below needs to address them directly.
Private Const COL_SORT_KEY As Long = 23
Private Const COL_CAMPAIGN_SORT_HELPER As Long = 24
Private Const COL_CAMPAIGN_RANK As Long = 25
Private Const COL_CAMPAIGN_RANK_KEY As Long = 26
Private Const COL_FRIENDLY_TIME_SPENT As Long = 27

Private Const FIRST_DATA_ROW As Long = 2

' --- Campaigns column layout (sheets/campaigns.py) -------------------------
Private Const CAMPAIGNS_FIRST_MAIN_ROW As Long = 3
Private Const CAMPAIGNS_SLOT_SIZE As Long = 10
Private Const CAMPAIGNS_NUM_SLOTS As Long = 20
Private Const CAMPAIGNS_COL_NAME As Long = 3
Private Const CAMPAIGNS_COL_KEY As Long = 25

' --- Campaigns Archive column layout (sheets/campaigns_archive.py) ---------
Private Const CAMPAIGNS_ARCHIVE_COL_NAME As Long = 1
Private Const CAMPAIGNS_ARCHIVE_COL_KEY As Long = 8

' --- Row extender (see ExtendTaskRows below) --------------------------------
Private Const ROW_EXTENDER_INCREMENT As Long = 500

' Public so Recurring Tasks' generation macro (modRecurringTasks.bas) can
' reuse it directly when writing into Tasks -- avoids a second,
' divergence-prone copy of Tasks' own column layout.
Public Function TasksColMap() As HierarchyColumnMap
    Dim m As HierarchyColumnMap
    m.FirstDataRow = FIRST_DATA_ROW
    m.LastCol = COL_LINKED_CAMPAIGN_KEY
    m.ColID = COL_ID
    m.ColTask = COL_TASK
    m.ColNumSubtasks = COL_NUM_SUBTASKS
    m.ColDaysLeft = COL_DAYS_LEFT
    m.ColSubtaskOf = COL_SUBTASK_OF
    m.ColKey = COL_TASK_KEY
    m.ColParentKey = COL_PARENT_KEY
    m.ColLevel = COL_LEVEL
    m.ColSiblingOrder = COL_SIBLING_ORDER
    TasksColMap = m
End Function

' ============================================================================
' Worksheet_Change handlers
' ============================================================================

' Also the bootstrap point for a genuinely NEW task row: §6.2 point 1 says
' every row "gets a hidden, never-reused TaskKey... assigned once at
' creation", but the spec text doesn't say *when* -- this is it. A row
' with no TaskKey yet is, by construction, brand new (TaskKey is never
' cleared once set), so a first non-blank Task value both stamps Start
' Date AND assigns TaskKey/SiblingOrder/an initial RenumberHierarchy pass
' (root-level, until the user sets Subtask of), so the task gets a real
' ID immediately instead of staying blank until some unrelated structural
' change happens to trigger a renumber.
Public Function HandleTaskNameChange(r As Long) As String
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")

    If Len(Trim$(ws.Cells(r, COL_TASK).Value & "")) > 0 Then
        Dim isNewRow As Boolean
        isNewRow = (Len(Trim$(ws.Cells(r, COL_TASK_KEY).Value & "")) = 0)

        If Len(Trim$(ws.Cells(r, COL_START_DATE).Value & "")) = 0 Then
            ws.Cells(r, COL_START_DATE).Value = Now()
        End If

        ' User-requested default (PLAN.md §19 Phase 5): a brand-new task
        ' starts life as "Not Started" rather than sitting blank until the
        ' user manually picks a Status.
        If Len(Trim$(ws.Cells(r, COL_STATUS).Value & "")) = 0 Then
            ws.Cells(r, COL_STATUS).Value = "Not Started"
        End If

        ' User-requested default (PLAN.md §19 Phase 5): a brand-new task
        ' starts life Linked to "None" rather than sitting blank -- "None"
        ' is a permanent, literal first entry in the Linked Campaign
        ' dropdown's source list (sheets/campaigns.py's CampaignNameList,
        ' Campaigns!X3), not a real campaign, so no downstream lookup
        ' (HandleLinkedCampaignChange) ever matches it -- LinkedCampaignKey
        ' naturally stays blank, identical to "genuinely unlinked."
        If Len(Trim$(ws.Cells(r, COL_LINKED_CAMPAIGN).Value & "")) = 0 Then
            ws.Cells(r, COL_LINKED_CAMPAIGN).Value = "None"
        End If

        If isNewRow Then
            ws.Cells(r, COL_TASK_KEY).Value = NextTaskKey()
            Dim colMap As HierarchyColumnMap
            colMap = TasksColMap()
            ws.Cells(r, COL_SIBLING_ORDER).Value = NextSiblingOrder(ws, colMap, "")

            ' User-requested defaults (PLAN.md §19 Phase 5): populate the
            ' Up/Down/Timer cells as soon as a task is created, instead of
            ' leaving them blank until double-clicked -- an all-white row
            ' with no visible affordance made it unclear those cells were
            ' even clickable. RenumberHierarchy snapshots every column's
            ' value per row (including these), so setting them here, before
            ' the row potentially moves during renumbering, is safe -- they
            ' travel with the row like any other cell.
            ws.Cells(r, COL_UP).Value = ICON_UP
            ws.Cells(r, COL_DOWN).Value = ICON_DOWN
            ws.Cells(r, COL_TIMER).Value = ICON_TIMER_RUNNING & " Start"
            ' 0, not blank (user-requested, PLAN.md §19 Phase 5) -- Time
            ' Spent (hrs) is hidden, but its value drives the visible
            ' friendly "Time Spent" column's formula (=IF(ts="","",...)),
            ' so a genuinely blank hrs cell showed as a blank friendly
            ' cell instead of "0h 00m".
            ws.Cells(r, COL_TIME_SPENT).Value = 0

            RenumberHierarchy ws, TasksColMap()
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

' Re-applies the Up/Down/Timer "button" look (§19 Phase 6, user-reported)
' to every row that currently holds a real task, after any operation that
' can add/remove/reorder rows. Deliberately NOT a fixed style baked in
' across the sheet's whole 500-row allowance at build time -- that made
' the button borders visibly continue downward through hundreds of empty
' rows with no task in them. Re-scanning to the current last real row
' (not tracking "where did row r end up") is also what makes this safe to
' call after RenumberHierarchy specifically: a row's physical position
' can change during renumbering, but this always targets wherever the
' data actually landed, not a stale row number captured before the call.
Public Sub RefreshButtonStyling(ws As Worksheet)
    Dim lastRow As Long, r As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)
    For r = FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, COL_TASK_KEY).Value & "")) > 0 Then
            modDesignSystem.ApplyButtonLook ws.Cells(r, COL_UP)
            modDesignSystem.ApplyButtonLook ws.Cells(r, COL_DOWN)
            modDesignSystem.ApplyButtonLook ws.Cells(r, COL_TIMER)
        End If
    Next r
End Sub

' Computes the next monotonic TaskKey. Must check both Tasks (live) AND
' Archive (frozen) -- an archived task's key must never collide with a
' new one, and only checking Tasks' current MAX would miss that. Public
' so Recurring Tasks' generation macro can call it directly for each new
' instance row it writes into Tasks (single source of truth for "what's
' the next TaskKey" -- see TasksColMap's own Public note above).
Public Function NextTaskKey() As Double
    Dim ws As Worksheet, archiveWs As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")
    Set archiveWs = ThisWorkbook.Sheets("Archive")

    Dim maxKey As Double
    maxKey = 0

    Dim lastRow As Long, r As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)
    For r = FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, COL_TASK_KEY).Value & "")) > 0 Then
            If CDbl(ws.Cells(r, COL_TASK_KEY).Value) > maxKey Then maxKey = CDbl(ws.Cells(r, COL_TASK_KEY).Value)
        End If
    Next r

    Const ARCHIVE_COL_TASK_KEY As Long = 15
    Dim archLastRow As Long
    archLastRow = LastHierarchyRow(archiveWs, ARCHIVE_COL_TASK_KEY, 2)
    For r = 2 To archLastRow
        If Len(Trim$(archiveWs.Cells(r, ARCHIVE_COL_TASK_KEY).Value & "")) > 0 Then
            If CDbl(archiveWs.Cells(r, ARCHIVE_COL_TASK_KEY).Value) > maxKey Then maxKey = CDbl(archiveWs.Cells(r, ARCHIVE_COL_TASK_KEY).Value)
        End If
    Next r

    NextTaskKey = maxKey + 1
End Function

Public Function HandleSubtaskOfChange(r As Long) As String
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")

    Dim typedID As String
    typedID = Trim$(ws.Cells(r, COL_SUBTASK_OF).Value & "")

    Dim movingTaskKey As Double
    movingTaskKey = CDbl(ws.Cells(r, COL_TASK_KEY).Value)

    Dim newParentKey As String
    newParentKey = ""
    If Len(typedID) > 0 Then
        Dim foundRow As Long
        foundRow = FindRowByID(ws, typedID)
        If foundRow > 0 And foundRow <> r Then
            newParentKey = CStr(ws.Cells(foundRow, COL_TASK_KEY).Value)
        End If
    End If

    If Len(newParentKey) > 0 Then
        If WouldCreateCycle(ws, movingTaskKey, newParentKey) Then
            Application.EnableEvents = prevEvents
            MsgBox "That would make this task a subtask of itself (directly or indirectly). Change not applied.", vbExclamation, "Invalid Subtask Of"
            Application.EnableEvents = False
            RenumberHierarchy ws, TasksColMap()   ' restores the displayed Subtask of text to its real current value
            RefreshButtonStyling ws
            Application.EnableEvents = prevEvents
            HandleSubtaskOfChange = "REJECTED_CYCLE"
            Exit Function
        End If
    End If

    Dim currentParentKey As String
    currentParentKey = Trim$(ws.Cells(r, COL_PARENT_KEY).Value & "")

    If newParentKey <> currentParentKey Then
        ws.Cells(r, COL_PARENT_KEY).Value = newParentKey
        Dim colMap As HierarchyColumnMap
        colMap = TasksColMap()
        ws.Cells(r, COL_SIBLING_ORDER).Value = NextSiblingOrder(ws, colMap, newParentKey)
    End If

    RenumberHierarchy ws, TasksColMap()
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
    Set ws = ThisWorkbook.Sheets("Tasks")

    Dim typedName As String
    typedName = Trim$(ws.Cells(r, COL_LINKED_CAMPAIGN).Value & "")

    If Len(typedName) = 0 Then
        ws.Cells(r, COL_LINKED_CAMPAIGN_KEY).Value = ""
    Else
        Dim campaignsWs As Worksheet
        Set campaignsWs = ThisWorkbook.Sheets("Campaigns")
        Dim found As Boolean
        found = False
        Dim slot As Long, slotRow As Long
        For slot = 0 To CAMPAIGNS_NUM_SLOTS - 1
            slotRow = CAMPAIGNS_FIRST_MAIN_ROW + slot * CAMPAIGNS_SLOT_SIZE
            If Trim$(campaignsWs.Cells(slotRow, CAMPAIGNS_COL_NAME).Value & "") = typedName Then
                ws.Cells(r, COL_LINKED_CAMPAIGN_KEY).Value = campaignsWs.Cells(slotRow, CAMPAIGNS_COL_KEY).Value
                found = True
                Exit For
            End If
        Next slot
        If Not found Then
            ws.Cells(r, COL_LINKED_CAMPAIGN_KEY).Value = ""
        End If
    End If

    Application.EnableEvents = prevEvents
    HandleLinkedCampaignChange = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandleLinkedCampaignChange = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Refreshes every Task row's displayed Linked Campaign text to match its
' hidden LinkedCampaignKey's *current* name -- same "resolve by key, keep
' the label refreshed" pattern as Subtask of/ParentKey (§6.2 point 5).
' Unlike Subtask of (refreshed as a side effect of RenumberHierarchy's own
' column-map loop), LinkedCampaignKey isn't part of modHierarchy's generic
' column map, so this is a separate pass, called after every
' RenumberHierarchy plus directly after a Linked Campaign edit. Looks up
' live Campaigns first, then Campaigns Archive, so a task linked to an
' already-completed campaign still shows the right name (§6.2 point 5).
Public Sub RefreshLinkedCampaignLabels(ws As Worksheet)
    Dim campaignsWs As Worksheet, campaignsArchiveWs As Worksheet
    Set campaignsWs = ThisWorkbook.Sheets("Campaigns")
    Set campaignsArchiveWs = ThisWorkbook.Sheets("Campaigns Archive")

    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    Dim r As Long
    For r = FIRST_DATA_ROW To lastRow
        Dim keyText As String
        keyText = Trim$(ws.Cells(r, COL_LINKED_CAMPAIGN_KEY).Value & "")
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
                ws.Cells(r, COL_LINKED_CAMPAIGN).Value = campaignName
            End If
        End If
    Next r
End Sub

Public Function HandleStatusChange(r As Long) As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")

    Dim newStatus As String
    newStatus = Trim$(ws.Cells(r, COL_STATUS).Value & "")

    Select Case newStatus
        Case "Completed"
            CascadeComplete ws, r
            HandleStatusChange = "COMPLETED"
        Case "Delete"
            Dim confirmResult As VbMsgBoxResult
            confirmResult = MsgBox("Delete this task and all its subtasks? This cannot be undone.", vbYesNo + vbExclamation, "Confirm Delete")
            If confirmResult = vbYes Then
                CascadeDelete ws, r
                HandleStatusChange = "DELETED"
            Else
                Dim prevEvents As Boolean
                prevEvents = Application.EnableEvents
                Application.EnableEvents = False
                ws.Cells(r, COL_STATUS).Value = ""   ' can't know the true prior status -- clear rather than guess
                Application.EnableEvents = prevEvents
                HandleStatusChange = "DELETE_CANCELLED"
            End If
        Case Else
            HandleStatusChange = "OK"
    End Select
    Exit Function
ErrHandler:
    HandleStatusChange = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Collects a task row and every descendant row (ID-prefix match, same
' technique as the Number of Subtasks formula) into outRows/outCount.
Private Sub FindBranchRows(ws As Worksheet, r As Long, ByRef outRows() As Long, ByRef outCount As Long)
    Dim idText As String
    idText = Trim$(ws.Cells(r, COL_ID).Value & "")

    outCount = 0
    ReDim outRows(1 To 1000)
    outCount = outCount + 1
    outRows(outCount) = r

    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)
    Dim i As Long
    For i = FIRST_DATA_ROW To lastRow
        If i <> r Then
            Dim otherID As String
            otherID = Trim$(ws.Cells(i, COL_ID).Value & "")
            If Left$(otherID, Len(idText) + 1) = idText & "." Then
                outCount = outCount + 1
                outRows(outCount) = i
            End If
        End If
    Next i
End Sub

Private Sub CascadeComplete(ws As Worksheet, r As Long)
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim rowsInBranch() As Long
    Dim n As Long
    FindBranchRows ws, r, rowsInBranch, n

    Dim completedAt As Date
    completedAt = Now()

    Dim archiveWs As Worksheet
    Set archiveWs = ThisWorkbook.Sheets("Archive")
    archiveWs.Unprotect

    Dim archTbl As ListObject
    Set archTbl = archiveWs.ListObjects("ArchiveTable")

    SortAscending rowsInBranch, n   ' archive in visual (depth-first) order
    Dim k As Long
    For k = 1 To n
        ArchiveOneRow ws, rowsInBranch(k), archTbl, completedAt
    Next k

    archiveWs.Protect UserInterfaceOnly:=True

    SortDescending rowsInBranch, n   ' delete bottom-up so row indexes stay valid
    For k = 1 To n
        ws.Rows(rowsInBranch(k)).Delete
    Next k

    RenumberHierarchy ws, TasksColMap()
    RefreshButtonStyling ws
    RefreshLinkedCampaignLabels ws

    Application.EnableEvents = prevEvents
End Sub

' Public (not Private) specifically so it's callable directly via
' Application.Run during testing, bypassing HandleStatusChange's MsgBox
' confirmation -- a blocking modal can't be driven from automated COM
' testing (same reasoning as frmNotes' ShowForCell/PrepareForCell split).
' The production path (a real user setting Status="Delete") still goes
' through HandleStatusChange's confirm-then-cascade flow unaffected.
Public Sub CascadeDelete(ws As Worksheet, r As Long)
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim rowsInBranch() As Long
    Dim n As Long
    FindBranchRows ws, r, rowsInBranch, n

    SortDescending rowsInBranch, n
    Dim k As Long
    For k = 1 To n
        ws.Rows(rowsInBranch(k)).Delete
    Next k

    RenumberHierarchy ws, TasksColMap()
    RefreshButtonStyling ws
    RefreshLinkedCampaignLabels ws

    Application.EnableEvents = prevEvents
    Exit Sub
ErrHandler:
    Application.EnableEvents = prevEvents
    Err.Raise Err.Number, , Err.Description
End Sub

' Column order matches archive.py's HEADERS + HIDDEN_HEADERS exactly
' Archive's table starts at the "minimum valid size" (header + one blank
' data row, sheets/archive.py's own docstring) rather than zero data rows
' -- so a naive ListRows.Add always appends AFTER that still-blank
' template row, leaving it permanently empty and pushing every real
' archived row one down, which reads as a stray gap between the header
' and the first real entry (confirmed live, PLAN.md §19 Phase 5). Reuses
' that first row directly the one time it's still genuinely blank
' (Archive is protected against any other kind of edit, so "exactly one
' row, and it's blank" can only mean "nothing has ever been archived
' yet" -- never a real archived row that happens to look empty).
Private Function NextArchiveRow(archTbl As ListObject) As ListRow
    If archTbl.ListRows.Count = 1 Then
        Dim firstRow As ListRow
        Set firstRow = archTbl.ListRows(1)
        If Len(Trim$(firstRow.Range(1, 1).Value & "")) = 0 Then
            Set NextArchiveRow = firstRow
            Exit Function
        End If
    End If
    Set NextArchiveRow = archTbl.ListRows.Add
End Function

' (PLAN.md §7.1).
Private Sub ArchiveOneRow(tasksWs As Worksheet, taskRow As Long, archTbl As ListObject, completedAt As Date)
    Dim newRow As ListRow
    Set newRow = NextArchiveRow(archTbl)

    newRow.Range(1, 1).Value = tasksWs.Cells(taskRow, COL_ID).Value
    newRow.Range(1, 2).Value = tasksWs.Cells(taskRow, COL_START_DATE).Value
    newRow.Range(1, 3).Value = tasksWs.Cells(taskRow, COL_TASK).Value
    newRow.Range(1, 4).Value = tasksWs.Cells(taskRow, COL_NUM_SUBTASKS).Value
    newRow.Range(1, 5).Value = tasksWs.Cells(taskRow, COL_SUBTASK_OF).Value
    newRow.Range(1, 6).Value = tasksWs.Cells(taskRow, COL_PROJECT).Value
    newRow.Range(1, 7).Value = tasksWs.Cells(taskRow, COL_LINKED_CAMPAIGN).Value
    newRow.Range(1, 8).Value = tasksWs.Cells(taskRow, COL_PRIORITY).Value
    newRow.Range(1, 9).Value = tasksWs.Cells(taskRow, COL_DEADLINE).Value
    newRow.Range(1, 10).Value = "Completed"

    Dim deadlineVal As Variant
    deadlineVal = tasksWs.Cells(taskRow, COL_DEADLINE).Value
    If Len(Trim$(deadlineVal & "")) > 0 Then
        newRow.Range(1, 11).Value = CDate(deadlineVal) - completedAt   ' frozen "days early/late", §7.1
    End If

    newRow.Range(1, 12).Value = tasksWs.Cells(taskRow, COL_TIME_SPENT).Value
    newRow.Range(1, 13).Value = tasksWs.Cells(taskRow, COL_NOTES).Value
    newRow.Range(1, 14).Value = completedAt

    newRow.Range(1, 15).Value = tasksWs.Cells(taskRow, COL_TASK_KEY).Value
    newRow.Range(1, 16).Value = tasksWs.Cells(taskRow, COL_PARENT_KEY).Value
    newRow.Range(1, 17).Value = tasksWs.Cells(taskRow, COL_LEVEL).Value
End Sub

' ============================================================================
' Worksheet_BeforeDoubleClick handlers
' ============================================================================

Public Function HandleDeadlineDoubleClick(r As Long) As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")
    Dim taskKey As String
    taskKey = CStr(ws.Cells(r, COL_TASK_KEY).Value)

    SetCalendarHandoff "Tasks", taskKey, "TaskKey", "Deadline"
    ThisWorkbook.Sheets("Calendar").Activate

    HandleDeadlineDoubleClick = "OK"
    Exit Function
ErrHandler:
    HandleDeadlineDoubleClick = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function HandleTimerDoubleClick(r As Long) As String
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")

    Dim thisRunningSince As String
    thisRunningSince = Trim$(ws.Cells(r, COL_TIMER_RUNNING_SINCE).Value & "")

    ' Only one timer can run at a time (§6, confirmed decision) -- stop any
    ' OTHER row's running timer first.
    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)
    Dim i As Long
    For i = FIRST_DATA_ROW To lastRow
        If i <> r Then
            If Len(Trim$(ws.Cells(i, COL_TIMER_RUNNING_SINCE).Value & "")) > 0 Then
                StopTimer ws, i
            End If
        End If
    Next i

    If Len(thisRunningSince) = 0 Then
        StartTimer ws, r
        HandleTimerDoubleClick = "STARTED"
    Else
        StopTimer ws, r
        HandleTimerDoubleClick = "STOPPED"
    End If

    Application.EnableEvents = prevEvents
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandleTimerDoubleClick = "ERROR " & Err.Number & ": " & Err.Description
End Function

Private Sub StartTimer(ws As Worksheet, r As Long)
    ws.Cells(r, COL_TIMER_RUNNING_SINCE).Value = Now()
    Dim timerCell As Range
    Set timerCell = ws.Cells(r, COL_TIMER)
    timerCell.Value = ICON_TIMER_RUNNING & " Running"
    timerCell.Interior.Color = COLOR_STATUS_GOOD

    Dim currentStatus As String
    currentStatus = Trim$(ws.Cells(r, COL_STATUS).Value & "")
    If currentStatus <> "In Progress" And currentStatus <> "Completed" Then
        ws.Cells(r, COL_STATUS).Value = "In Progress"
    End If
End Sub

Private Sub StopTimer(ws As Worksheet, r As Long)
    Dim startedAt As Date
    startedAt = CDate(ws.Cells(r, COL_TIMER_RUNNING_SINCE).Value)
    Dim elapsedHours As Double
    elapsedHours = (Now() - startedAt) * 24

    Dim currentSpent As Double
    currentSpent = 0
    If Len(Trim$(ws.Cells(r, COL_TIME_SPENT).Value & "")) > 0 Then
        currentSpent = CDbl(ws.Cells(r, COL_TIME_SPENT).Value)
    End If
    ws.Cells(r, COL_TIME_SPENT).Value = currentSpent + elapsedHours

    ws.Cells(r, COL_TIMER_RUNNING_SINCE).Value = ""
    Dim timerCell As Range
    Set timerCell = ws.Cells(r, COL_TIMER)
    ' Resets to the same "Start" affordance shown on a brand-new task
    ' (see HandleTaskNameChange), not blank -- so the cell always reads
    ' as a clickable button, never as empty/dead space.
    timerCell.Value = ICON_TIMER_RUNNING & " Start"
    timerCell.Interior.Pattern = xlNone

    LogTimeEntry ws.Cells(r, COL_TASK_KEY).Value, VBA.Date, elapsedHours
End Sub

Private Sub LogTimeEntry(taskKey As Variant, entryDate As Date, hours As Double)
    Dim tlWs As Worksheet
    Set tlWs = ThisWorkbook.Sheets("Time Log")
    Dim tlTbl As ListObject
    Set tlTbl = tlWs.ListObjects("TimeLogTable")
    Dim newRow As ListRow
    Set newRow = tlTbl.ListRows.Add
    newRow.Range(1, 1).Value = taskKey
    newRow.Range(1, 2).Value = entryDate
    newRow.Range(1, 3).Value = hours
End Sub

Public Function HandleTaskCollapseDoubleClick(r As Long) As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")

    Dim numSubtasks As Long
    numSubtasks = 0
    If Len(Trim$(ws.Cells(r, COL_NUM_SUBTASKS).Value & "")) > 0 Then
        numSubtasks = CLng(ws.Cells(r, COL_NUM_SUBTASKS).Value)
    End If

    If numSubtasks = 0 Then
        HandleTaskCollapseDoubleClick = "NO_SUBTASKS"
        Exit Function
    End If

    Dim myLevel As Long
    myLevel = CLng(ws.Cells(r, COL_LEVEL).Value)

    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)

    Dim blockEnd As Long
    blockEnd = r
    Dim i As Long
    For i = r + 1 To lastRow
        If CLng(ws.Cells(i, COL_LEVEL).Value) > myLevel Then
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
    Set ws = ThisWorkbook.Sheets("Tasks")

    Dim myParentKey As String
    myParentKey = Trim$(ws.Cells(r, COL_PARENT_KEY).Value & "")
    Dim myLevel As Long
    myLevel = CLng(ws.Cells(r, COL_LEVEL).Value)

    Dim targetRow As Long
    targetRow = 0

    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)

    If LCase$(direction) = "up" Then
        Dim i As Long
        For i = r - 1 To FIRST_DATA_ROW Step -1
            If CLng(ws.Cells(i, COL_LEVEL).Value) < myLevel Then Exit For
            If CLng(ws.Cells(i, COL_LEVEL).Value) = myLevel And Trim$(ws.Cells(i, COL_PARENT_KEY).Value & "") = myParentKey Then
                targetRow = i
                Exit For
            End If
        Next i
    Else
        Dim j As Long
        For j = r + 1 To lastRow
            If CLng(ws.Cells(j, COL_LEVEL).Value) < myLevel Then Exit For
            If CLng(ws.Cells(j, COL_LEVEL).Value) = myLevel And Trim$(ws.Cells(j, COL_PARENT_KEY).Value & "") = myParentKey Then
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
    myOrder = CDbl(ws.Cells(r, COL_SIBLING_ORDER).Value)
    targetOrder = CDbl(ws.Cells(targetRow, COL_SIBLING_ORDER).Value)

    ws.Cells(r, COL_SIBLING_ORDER).Value = targetOrder
    ws.Cells(targetRow, COL_SIBLING_ORDER).Value = myOrder

    RenumberHierarchy ws, TasksColMap()
    RefreshButtonStyling ws
    RefreshLinkedCampaignLabels ws

    Application.EnableEvents = prevEvents
    HandleSiblingReorder = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandleSiblingReorder = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Called by ThisWorkbook.bas's Workbook_Open (Phase 3, later item) to warn
' if a timer was left running when the file was last closed.
Public Function FindRunningTimerTaskName() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")
    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)
    Dim r As Long
    For r = FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, COL_TIMER_RUNNING_SINCE).Value & "")) > 0 Then
            FindRunningTimerTaskName = ws.Cells(r, COL_TASK).Value
            Exit Function
        End If
    Next r
    FindRunningTimerTaskName = ""
    Exit Function
ErrHandler:
    FindRunningTimerTaskName = ""
End Function

' ============================================================================
' Small helpers
' ============================================================================

Private Function FindRowByID(ws As Worksheet, idText As String) As Long
    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)
    Dim r As Long
    For r = FIRST_DATA_ROW To lastRow
        If Trim$(ws.Cells(r, COL_ID).Value & "") = idText Then
            FindRowByID = r
            Exit Function
        End If
    Next r
    FindRowByID = 0
End Function

Private Function FindRowByTaskKey(ws As Worksheet, taskKey As Double) As Long
    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)
    Dim r As Long
    For r = FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, COL_TASK_KEY).Value & "")) > 0 Then
            If CDbl(ws.Cells(r, COL_TASK_KEY).Value) = taskKey Then
                FindRowByTaskKey = r
                Exit Function
            End If
        End If
    Next r
    FindRowByTaskKey = 0
End Function

' Walks proposedParentKey's own ancestor chain looking for movingTaskKey --
' if found, re-parenting would create a cycle (RenumberHierarchy's tree
' walk only reaches nodes connected to root; a cycle would silently drop
' those rows from the rewrite, losing their data). guard caps the walk at
' 10000 steps as a final backstop against an already-corrupt chain.
' Public (not Private) specifically so the detection logic is testable
' directly via Application.Run -- HandleSubtaskOfChange's own cycle path
' calls MsgBox, a blocking modal that must never fire during automated
' testing (same reasoning as CascadeDelete's Public visibility above).
Public Function WouldCreateCycle(ws As Worksheet, movingTaskKey As Double, proposedParentKey As String) As Boolean
    Dim currentKey As String
    currentKey = proposedParentKey
    Dim guard As Long
    guard = 0
    Do While Len(currentKey) > 0 And guard < 10000
        If CDbl(currentKey) = movingTaskKey Then
            WouldCreateCycle = True
            Exit Function
        End If
        Dim parentRow As Long
        parentRow = FindRowByTaskKey(ws, CDbl(currentKey))
        If parentRow = 0 Then Exit Do
        currentKey = Trim$(ws.Cells(parentRow, COL_PARENT_KEY).Value & "")
        guard = guard + 1
    Loop
    WouldCreateCycle = False
End Function

Private Sub SortAscending(arr() As Long, n As Long)
    Dim i As Long, j As Long, tmp As Long
    For i = 2 To n
        tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j) > tmp Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = tmp
    Next i
End Sub

Private Sub SortDescending(arr() As Long, n As Long)
    SortAscending arr, n
    Dim i As Long, j As Long, tmp As Long
    i = 1: j = n
    Do While i < j
        tmp = arr(i)
        arr(i) = arr(j)
        arr(j) = tmp
        i = i + 1
        j = j - 1
    Loop
End Sub

' ============================================================================
' Row extender -- "keep this tracker usable for years" (user-requested):
' a permanent, re-runnable "add 500 more rows" action instead of a one-off
' manual row-count bump (PLAN.md §20 Bug #9's original fix). See
' sheets/tasks.py's own note next to TasksRowAllowance for the full design
' rationale.
'
' Every place Tasks' current row bound is baked into something has to be
' touched on every extension -- found by an exhaustive audit, not just
' patching whatever Bug #9's fix happened to touch:
'   - Tasks' own Number of Subtasks (col F) and CampaignRank (col Y)
'     formulas -- each embeds an absolute $2:$<bound> range INSIDE its
'     formula text, for EVERY row (not just new ones), since
'     RenumberHierarchy deliberately never rewrites either column (see
'     COL_SORT_KEY's comment above) -- a row's formula, once written,
'     keeps whatever bound it was written with forever.
'   - Dashboard's and Campaigns' cross-sheet SUMIFS/COUNTIFS/lookup
'     ranges -- also literal $2:$<bound> text, but scattered across many
'     formulas per sheet rather than one column, so these are patched via
'     a sheet-wide literal-substring Replace instead of enumerating cells.
'   - Data validation and conditional formatting on the new rows.
'   - modCampaigns.bas's own TASKS_LAST_DATA_ROW bound -- now a Function
'     reading TasksRowAllowance directly, so it no longer needs touching.
' ============================================================================

' TasksRowAllowance is a workbook-scoped Defined Name holding a bare
' constant (sheets/tasks.py's build-time bootstrap), not a cell/range
' reference -- Evaluate() returns its numeric value directly either way.
' On Error GoTo, not just a bare read: this is called from
' RowExtenderBannerRow, which Tasks.bas's dispatcher calls on EVERY
' double-click anywhere on the sheet -- if the Name doesn't exist yet
' (workbook never bootstrapped) this would otherwise throw unhandled on
' every single double-click, popping CLAUDE.md's documented blocking
' modal. Returns 0 on failure so RowExtenderBannerRow degrades to "row 1"
' (never a real Target.Row -- Tasks.bas's dispatcher already exits before
' row 2) instead of propagating an error.
Private Function TasksRowAllowance() As Long
    On Error GoTo ErrHandler
    TasksRowAllowance = CLng(Application.Evaluate("TasksRowAllowance"))
    Exit Function
ErrHandler:
    TasksRowAllowance = 0
End Function

Private Sub SetTasksRowAllowance(ByVal n As Long)
    ThisWorkbook.Names("TasksRowAllowance").RefersTo = "=" & n
End Sub

' Public so Tasks.bas's dispatcher can check "did the user just double-
' click the banner" without duplicating the +1 math.
Public Function RowExtenderBannerRow() As Long
    RowExtenderBannerRow = TasksRowAllowance() + 1
End Function

' One-time bootstrap for a workbook built BEFORE this feature existed (the
' user's real live laptop file) -- a workbook freshly built via
' build_dashboard.py + inject_vba.py already has both the Name and the
' banner from sheets/tasks.py's build() step, so this is a no-op there.
' Safe to re-run: does nothing (just reports current state) if
' TasksRowAllowance already exists.
'
' Split into a MsgBox-free core (independently callable via Application.Run,
' same convention as ExtendTaskRows/HandleRowExtenderClick below) and a thin
' interactive wrapper -- BootstrapRowExtender itself must never be driven
' via Application.Run in a COM test, since its own MsgBox would block
' waiting for a click exactly like an unhandled runtime error does.
Public Function BootstrapRowExtenderCore() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")

    Dim already As Boolean
    already = False
    Dim nm As Variant
    For Each nm In ThisWorkbook.Names
        If nm.Name = "TasksRowAllowance" Then already = True
    Next nm

    If already Then
        BootstrapRowExtenderCore = "ALREADY (" & TasksRowAllowance() & ")"
        Exit Function
    End If

    ThisWorkbook.Names.Add Name:="TasksRowAllowance", RefersTo:="=501"
    DrawRowExtenderBanner ws, 501
    BootstrapRowExtenderCore = "OK"
    Exit Function
ErrHandler:
    BootstrapRowExtenderCore = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Sub BootstrapRowExtender()
    Dim result As String
    result = BootstrapRowExtenderCore()
    If result = "OK" Then
        MsgBox "Row extender is ready. A ""+ Add 500 More Task Rows"" button now appears at row 502.", _
            vbInformation, "Row extender bootstrap"
    ElseIf Left(result, 7) = "ALREADY" Then
        MsgBox "TasksRowAllowance already exists. Nothing to do.", vbInformation, "Row extender bootstrap"
    Else
        MsgBox "Bootstrap failed:" & vbNewLine & result, vbCritical, "Row extender bootstrap"
    End If
End Sub

Public Sub HandleRowExtenderClick()
    Dim result As String
    result = ExtendTaskRows()
    If result = "OK" Then
        MsgBox "Added " & ROW_EXTENDER_INCREMENT & " more rows. Tasks now goes up to row " & _
            TasksRowAllowance() & ".", vbInformation, "Rows added"
    Else
        MsgBox "Couldn't add more rows:" & vbNewLine & result, vbExclamation, "Row extender"
    End If
End Sub

' Core logic, independently callable via Application.Run (bypasses the
' MsgBox above -- same split as every other MsgBox-fronted handler in this
' module). Returns "OK" or "ERROR <n>: <description>", never lets an error
' propagate unhandled.
Public Function ExtendTaskRows() As String
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean, prevScreen As Boolean
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim ws As Worksheet, dashWs As Worksheet, campWs As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")
    Set dashWs = ThisWorkbook.Sheets("Dashboard")
    Set campWs = ThisWorkbook.Sheets("Campaigns")

    Dim oldLast As Long, newLast As Long
    oldLast = TasksRowAllowance()
    newLast = oldLast + ROW_EXTENDER_INCREMENT

    ' The banner cell (A<oldLast+1>:F<oldLast+1>) is about to become an
    ' ordinary data row -- unmerge and strip its special look before the
    ' template-row loop below formats it like every other new row.
    Dim bannerRange As Range
    Set bannerRange = ws.Range(ws.Cells(oldLast + 1, 1), ws.Cells(oldLast + 1, 6))
    bannerRange.UnMerge
    bannerRange.ClearFormats
    bannerRange.ClearContents

    ' Built from scratch, never copied from an existing row -- PLAN.md §20
    ' Bug #9 attempt #1's exact mistake was copying the (possibly real,
    ' possibly button-styled) last row's format onto new blank rows.
    Dim r As Long
    For r = oldLast + 1 To newLast
        BuildTaskTemplateRow ws, r, newLast
    Next r

    ' Both range-bound formula columns get rewritten for ALL rows 2..newLast
    ' (old rows included) -- their embedded bound is stale everywhere, not
    ' just past the old boundary.
    RewriteNumSubtasksFormulas ws, newLast
    RewriteCampaignRankFormulas ws, newLast

    ExtendValidationAndCF ws, oldLast, newLast

    ReplaceRowBound dashWs, oldLast, newLast
    ReplaceRowBound campWs, oldLast, newLast

    SetTasksRowAllowance newLast
    DrawRowExtenderBanner ws, newLast

    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    ExtendTaskRows = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    ExtendTaskRows = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Builds one brand-new row's number formats, position-relative formulas,
' alignment, and hairline border -- mirrors sheets/tasks.py's build()
' loop exactly (kept in sync by hand, same convention as every other
' cross-file column-layout constant in this project). newLast is the
' CampaignRank/Number-of-Subtasks bound this row's formulas should use.
Private Sub BuildTaskTemplateRow(ws As Worksheet, r As Long, newLast As Long)
    Dim DQ As String
    DQ = Chr(34)

    ws.Cells(r, COL_START_DATE).NumberFormat = "DD.MM.YYYY HH:MM"
    ws.Cells(r, COL_DEADLINE).NumberFormat = "DD.MM.YYYY HH:MM"
    ws.Cells(r, COL_TIME_SPENT).NumberFormat = "0.00"
    ws.Cells(r, COL_NOTES).NumberFormat = "General;General;General;" & DQ & modDesignSystem.ICON_NOTES() & DQ

    Dim idRange As String, thisId As String, countExpr As String
    idRange = "$A$2:$A$" & newLast
    thisId = "A" & r
    countExpr = "COUNTIF(" & idRange & "," & thisId & "&" & DQ & ".*" & DQ & ")"
    ws.Cells(r, COL_NUM_SUBTASKS).Formula = "=IF(" & countExpr & "=0," & DQ & DQ & "," & countExpr & ")"

    Dim dl As String
    dl = "K" & r   ' Deadline
    ws.Cells(r, COL_DAYS_LEFT).Formula = "=IF(" & dl & "=" & DQ & DQ & "," & DQ & DQ & ",INT(" & dl & ")-TODAY())"

    Dim ts As String
    ts = "O" & r   ' Time Spent (hrs)
    ws.Cells(r, COL_FRIENDLY_TIME_SPENT).Formula = _
        "=IF(" & ts & "=" & DQ & DQ & "," & DQ & DQ & ",TEXT(INT(ROUND(" & ts & "*60,0)/60)," & DQ & "0" & DQ & ")&" & _
        DQ & "h " & DQ & "&TEXT(MOD(ROUND(" & ts & "*60,0),60)," & DQ & "00" & DQ & ")&" & DQ & "m" & DQ & ")"

    Dim key As String
    key = "Q" & r   ' TaskKey
    ws.Cells(r, COL_SORT_KEY).Formula = "=IF(" & dl & "=" & DQ & DQ & "," & DQ & DQ & "," & dl & "+" & key & "/1000000)"

    Dim lck As String
    lck = "V" & r   ' LinkedCampaignKey
    ws.Cells(r, COL_CAMPAIGN_SORT_HELPER).Formula = _
        "=IF(" & lck & "=" & DQ & DQ & "," & DQ & DQ & ",IF(" & dl & "=" & DQ & DQ & ",9999999+" & key & "/1000000," & _
        dl & "+" & key & "/1000000))"

    Dim cshRange As String, lckRange As String, csh As String
    cshRange = "$X$2:$X$" & newLast
    lckRange = "$V$2:$V$" & newLast
    csh = "X" & r
    ws.Cells(r, COL_CAMPAIGN_RANK).Formula = _
        "=IF(" & lck & "=" & DQ & DQ & "," & DQ & DQ & ",SUMPRODUCT((" & lckRange & "=" & lck & ")*(" & _
        cshRange & "<=" & csh & ")))"

    Dim rankCell As String
    rankCell = "Y" & r
    ws.Cells(r, COL_CAMPAIGN_RANK_KEY).Formula = _
        "=IF(" & lck & "=" & DQ & DQ & "," & DQ & DQ & "," & lck & "&" & DQ & "|" & DQ & "&" & rankCell & ")"

    Dim c As Variant
    For Each c In Array(COL_ID, COL_NUM_SUBTASKS, COL_SUBTASK_OF, COL_DAYS_LEFT, COL_FRIENDLY_TIME_SPENT)
        With ws.Cells(r, CLng(c))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_FRIENDLY_TIME_SPENT)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = modDesignSystem.COLOR_GRIDLINE
    End With
End Sub

' Rewrites Number of Subtasks (col F) for EVERY row 2..newLast -- old rows
' included, since their embedded $2:$<bound> range is stale too. Formula
' shape matches hierarchy.py's number_of_subtasks_formula() exactly.
Private Sub RewriteNumSubtasksFormulas(ws As Worksheet, newLast As Long)
    Dim DQ As String
    DQ = Chr(34)
    Dim idRange As String
    idRange = "$A$2:$A$" & newLast
    Dim r As Long, thisId As String, countExpr As String
    For r = FIRST_DATA_ROW To newLast
        thisId = "A" & r
        countExpr = "COUNTIF(" & idRange & "," & thisId & "&" & DQ & ".*" & DQ & ")"
        ws.Cells(r, COL_NUM_SUBTASKS).Formula = "=IF(" & countExpr & "=0," & DQ & DQ & "," & countExpr & ")"
    Next r
End Sub

' Rewrites CampaignRank (col Y) for EVERY row 2..newLast -- same reasoning
' as above. Formula shape matches sheets/tasks.py's build() loop exactly
' (PLAN.md §20 Bug #9's original fix, generalized to run on demand).
Private Sub RewriteCampaignRankFormulas(ws As Worksheet, newLast As Long)
    Dim DQ As String
    DQ = Chr(34)
    Dim cshRange As String, lckRange As String
    cshRange = "$X$2:$X$" & newLast
    lckRange = "$V$2:$V$" & newLast
    Dim r As Long, lck As String, csh As String
    For r = FIRST_DATA_ROW To newLast
        lck = "V" & r
        csh = "X" & r
        ws.Cells(r, COL_CAMPAIGN_RANK).Formula = _
            "=IF(" & lck & "=" & DQ & DQ & "," & DQ & DQ & ",SUMPRODUCT((" & lckRange & "=" & lck & ")*(" & _
            cshRange & "<=" & csh & ")))"
    Next r
End Sub

' Data validation + conditional formatting for the new rows. DV is added
' fresh only to the new range (old rows' DV is left completely alone). CF
' rules already applying to the old range are resized via
' ModifyAppliesToRange -- extends the SAME rule objects (preserving
' whatever fill/formula they already carry) rather than recreating them,
' sidestepping the openpyxl-vs-VBA CF-fill/border authoring quirks
' documented in CLAUDE.md entirely, since nothing is being newly created.
Private Sub ExtendValidationAndCF(ws As Worksheet, oldLast As Long, newLast As Long)
    Dim newRange As Range

    Set newRange = ws.Range(ws.Cells(oldLast + 1, COL_PRIORITY), ws.Cells(newLast, COL_PRIORITY))
    With newRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Low,Medium,High"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    Set newRange = ws.Range(ws.Cells(oldLast + 1, COL_STATUS), ws.Cells(newLast, COL_STATUS))
    With newRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Not Started,In Progress,Completed,Delete"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    Set newRange = ws.Range(ws.Cells(oldLast + 1, COL_LINKED_CAMPAIGN), ws.Cells(newLast, COL_LINKED_CAMPAIGN))
    With newRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=CampaignNameList"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    Set newRange = ws.Range(ws.Cells(oldLast + 1, COL_DEADLINE), ws.Cells(newLast, COL_DEADLINE))
    With newRange.Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=FALSE"
        .ErrorTitle = "Use the picker instead"
        .ErrorMessage = "Double-click this cell to pick a date from the Calendar instead."
    End With

    Set newRange = ws.Range(ws.Cells(oldLast + 1, COL_NOTES), ws.Cells(newLast, COL_NOTES))
    With newRange.Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=FALSE"
        .ErrorTitle = "Use the picker instead"
        .ErrorMessage = "Double-click this cell to open the Notes popup instead."
    End With

    ExtendCFAppliesTo ws, COL_PRIORITY, COL_PRIORITY, newLast
    ExtendCFAppliesTo ws, COL_DAYS_LEFT, COL_DAYS_LEFT, newLast
    ExtendCFAppliesTo ws, COL_STATUS, COL_STATUS, newLast
    ExtendCFAppliesTo ws, COL_TASK, COL_TASK, newLast                     ' subtask "prefix" rule
    ExtendCFAppliesTo ws, COL_NUM_SUBTASKS, COL_NUM_SUBTASKS, newLast     ' magnitude ramp
    ExtendCFAppliesTo ws, 1, 16, newLast                                  ' depth shading, A:P
End Sub

' Queries the SINGLE anchor cell (FIRST_DATA_ROW, firstCol) for its
' FormatConditions -- reliable per Excel's object model (unlike querying a
' large multi-row Range, which is not guaranteed to enumerate every rule
' that touches it) -- and resizes each one found to FIRST_DATA_ROW..newLast
' across firstCol..lastCol.
Private Sub ExtendCFAppliesTo(ws As Worksheet, firstCol As Long, lastCol As Long, newLast As Long)
    Dim newRange As Range
    Set newRange = ws.Range(ws.Cells(FIRST_DATA_ROW, firstCol), ws.Cells(newLast, lastCol))

    ' Direct iteration -- confirmed live via COM (scratchpad debug probes,
    ' this session) that snapshotting each FormatCondition into a
    ' `Dim conds() As FormatCondition` array first, then calling
    ' ModifyAppliesToRange on the array elements, throws "RPC_E_DISCONNECTED"
    ' on every single condition, on every column tried (Priority, DaysLeft,
    ' the 16-column depth-shading range). Storing a FormatCondition
    ' reference in an array element evidently invalidates its COM
    ' connection. Calling ModifyAppliesToRange directly on
    ' FormatConditions.Item(i) inside a plain For loop (no snapshot) works
    ' correctly and was verified not to hit any enumerator-invalidation
    ' issue in practice.
    Dim srcConditions As FormatConditions
    Set srcConditions = ws.Cells(FIRST_DATA_ROW, firstCol).FormatConditions
    Dim n As Long, i As Long
    n = srcConditions.Count
    For i = 1 To n
        srcConditions.Item(i).ModifyAppliesToRange newRange
    Next i
End Sub

' Sheet-wide literal-substring replace of the row bound, for sheets whose
' Tasks-referencing formulas are scattered across many cells rather than
' one column (Dashboard, Campaigns) -- same mechanism PLAN.md §20 Bug #9's
' original fix used, confirmed reliable there. Deliberately NO LookIn:=
' argument -- Range.Replace has no such parameter; passing one doesn't
' raise a compile error when injected via AddFromString, it hangs
' Application.Run indefinitely with no error (CLAUDE.md's hard-won
' lesson). LookAt:=xlPart is required since "$<bound>" is only ever part
' of a larger formula (e.g. "$K$2:$K$1500"), never the whole cell content.
Private Sub ReplaceRowBound(ws As Worksheet, oldLast As Long, newLast As Long)
    ws.Cells.Replace What:="$" & oldLast, Replacement:="$" & newLast, LookAt:=xlPart, MatchCase:=True
End Sub

' Moves the "add more rows" banner down to sit one row below the new
' allowance. Mirrors sheets/tasks.py's own banner exactly so a rebuilt
' file and a live-extended file always look identical. Uses BorderAround
' (not per-edge .Borders on the anchor cell alone) -- design_system.py's
' apply_box_outline docstring explains why the naive per-anchor-cell
' border renders incomplete on a merged range; BorderAround is VBA's own
' correct native equivalent for a merged cell.
Private Sub DrawRowExtenderBanner(ws As Worksheet, newLast As Long)
    Dim bannerRow As Long
    bannerRow = newLast + 1
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(bannerRow, 1), ws.Cells(bannerRow, 6))
    rng.Merge
    rng.Cells(1, 1).Value = "+ Add 500 More Task Rows"
    With rng
        .Font.Name = modDesignSystem.FONT_NAME
        .Font.Size = modDesignSystem.FONT_SIZE_BODY
        .Font.Bold = True
        .Font.Color = modDesignSystem.COLOR_INK_PRIMARY
        .Interior.Color = modDesignSystem.COLOR_CARD_SURFACE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=modDesignSystem.COLOR_INK_MUTED
    End With
End Sub
