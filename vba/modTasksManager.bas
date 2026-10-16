Attribute VB_Name = "modTasksManager"
Option Explicit

' Manager-variant Tasks business logic (PLAN.md, "Manager-variant workbook"
' §4). A genuine fork of modTasks.bas, not a runtime-branching shared
' module -- modTasks.bas is built entirely around bare column-index
' constants threaded through every function, and the manager layout
' renumbers nearly all of them (PIC/Completed Subtasks inserted, Timer/
' Time Spent removed), so branching one module at runtime would mean two
' parallel constant sets tangled in one file. The dispatcher is
' vba/TasksManager.bas (sheet code-behind for the Tasks worksheet in the
' manager build), mirroring vba/Tasks.bas's own thin-dispatcher pattern.
'
' Status="Delete" only ever touches Tasks and was fully implemented in
' Phase 4. Status="Completed" (this phase) additionally touches Archive --
' CascadeComplete, FindOrCreateArchiveRow, and the Completed Subtasks
' counter/ancestor-walk logic below implement plan §5/§6 against Archive's
' manager column layout (ID, Start Date, Task, PIC, Total Subtasks,
' Completed Subtasks, Subtask of, Project/Category, Linked Campaign,
' Priority, Deadline, Status, Days Left, Notes, Completed At, TaskKey,
' ParentKey, Level, RowType) -- fixed by the plan even though
' sheets/archive.py itself (the Python structural builder) isn't built
' until Phase 6, same "logic ahead of structure" approach Phase 4 used for
' Tasks' own manager layout.

' --- Manager Tasks column layout (sheets/tasks.py, variant="manager") -----
Private Const COL_ID As Long = 1
Private Const COL_UP As Long = 2
Private Const COL_DOWN As Long = 3
Private Const COL_START_DATE As Long = 4
Private Const COL_TASK As Long = 5
Private Const COL_PIC As Long = 6
Private Const COL_NUM_SUBTASKS As Long = 7          ' "Total Subtasks"
Private Const COL_COMPLETED_SUBTASKS As Long = 8
Private Const COL_SUBTASK_OF As Long = 9
Private Const COL_PROJECT As Long = 10
Private Const COL_LINKED_CAMPAIGN As Long = 11
Private Const COL_PRIORITY As Long = 12
Private Const COL_DEADLINE As Long = 13
Private Const COL_STATUS As Long = 14
Private Const COL_DAYS_LEFT As Long = 15
Private Const COL_NOTES As Long = 16
Private Const COL_TASK_KEY As Long = 17
Private Const COL_PARENT_KEY As Long = 18
Private Const COL_LEVEL As Long = 19
Private Const COL_SIBLING_ORDER As Long = 20
Private Const COL_LINKED_CAMPAIGN_KEY As Long = 21
' Columns 22-25 (SortKey/CampaignSortHelper/CampaignRank/CampaignRankKey)
' are excluded from TasksManagerColMap's LastCol below for the same
' self-healing-formula reason as modTasks.bas's identical exclusion (see
' its own comment) -- named here because the row extender needs them.
Private Const COL_SORT_KEY As Long = 22
Private Const COL_CAMPAIGN_SORT_HELPER As Long = 23
Private Const COL_CAMPAIGN_RANK As Long = 24
Private Const COL_CAMPAIGN_RANK_KEY As Long = 25

Private Const FIRST_DATA_ROW As Long = 2

' --- Campaigns column layout (sheets/campaigns.py) -- shared, unchanged ---
Private Const CAMPAIGNS_FIRST_MAIN_ROW As Long = 3
Private Const CAMPAIGNS_SLOT_SIZE As Long = 10
Private Const CAMPAIGNS_NUM_SLOTS As Long = 20
Private Const CAMPAIGNS_COL_NAME As Long = 3
Private Const CAMPAIGNS_COL_KEY As Long = 25

' --- Campaigns Archive column layout -- shared, unchanged ------------------
Private Const CAMPAIGNS_ARCHIVE_COL_NAME As Long = 1
Private Const CAMPAIGNS_ARCHIVE_COL_KEY As Long = 8

' NextTaskKey must also check Archive so a new TaskKey never collides with
' an already-archived one (same reasoning as modTasks.bas's NextTaskKey).
' 16 is TaskKey's position in the manager Archive layout documented in
' plan §6 (ID, Start Date, Task, PIC, Total Subtasks, Completed Subtasks,
' Subtask of, Project/Category, Linked Campaign, Priority, Deadline,
' Status, Days Left, Notes, Completed At, TaskKey, ...) -- that sheet
' itself isn't built until Phase 6, but the column position is already
' fixed by the plan, so it's safe to depend on here ahead of time.
Private Const ARCHIVE_COL_TASK_KEY As Long = 16

' --- Manager Archive column layout (sheets/archive.py, Phase 6 -- see
' ARCHIVE_COL_TASK_KEY's own comment for why it's safe to depend on this
' ahead of that sheet's own build) -----------------------------------------
Private Const ARCHIVE_COL_ID As Long = 1
Private Const ARCHIVE_COL_START_DATE As Long = 2
Private Const ARCHIVE_COL_TASK As Long = 3
Private Const ARCHIVE_COL_PIC As Long = 4
Private Const ARCHIVE_COL_TOTAL_SUBTASKS As Long = 5
Private Const ARCHIVE_COL_COMPLETED_SUBTASKS As Long = 6
Private Const ARCHIVE_COL_SUBTASK_OF As Long = 7
Private Const ARCHIVE_COL_PROJECT As Long = 8
Private Const ARCHIVE_COL_LINKED_CAMPAIGN As Long = 9
Private Const ARCHIVE_COL_PRIORITY As Long = 10
Private Const ARCHIVE_COL_DEADLINE As Long = 11
Private Const ARCHIVE_COL_STATUS As Long = 12
Private Const ARCHIVE_COL_DAYS_LEFT As Long = 13
Private Const ARCHIVE_COL_NOTES As Long = 14
Private Const ARCHIVE_COL_COMPLETED_AT As Long = 15
' ARCHIVE_COL_TASK_KEY = 16, already declared above
Private Const ARCHIVE_COL_PARENT_KEY As Long = 17
Private Const ARCHIVE_COL_LEVEL As Long = 18
Private Const ARCHIVE_COL_ROW_TYPE As Long = 19

' --- Row extender (see ExtendTaskRows below) --------------------------------
Private Const ROW_EXTENDER_INCREMENT As Long = 500

' --- PIC List column layout (sheets/pic_list.py) ---------------------------
Private Const PIC_FIRST_ROW As Long = 2
Private Const PIC_LAST_ROW As Long = 201
Private Const PIC_COL_NAME As Long = 1
Private Const PIC_COL_UP As Long = 2
Private Const PIC_COL_DOWN As Long = 3
Private Const PIC_COL_COLOR As Long = 4

Public Function TasksManagerColMap() As HierarchyColumnMap
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
    m.ColPIC = COL_PIC
    TasksManagerColMap = m
End Function

' ============================================================================
' Worksheet_Change handlers
' ============================================================================

' Bootstrap point for a genuinely new task row -- same trigger convention as
' modTasks.bas's HandleTaskNameChange (a row with no TaskKey yet is, by
' construction, brand new). No Timer/Time Spent defaults (manager tracks
' neither); PIC is deliberately left blank on creation, not defaulted --
' a new task starts in the unassigned group until the manager assigns it,
' per plan §2/§4.
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

        If Len(Trim$(ws.Cells(r, COL_STATUS).Value & "")) = 0 Then
            ws.Cells(r, COL_STATUS).Value = "Not Started"
        End If

        If Len(Trim$(ws.Cells(r, COL_LINKED_CAMPAIGN).Value & "")) = 0 Then
            ws.Cells(r, COL_LINKED_CAMPAIGN).Value = "None"
        End If

        If isNewRow Then
            ws.Cells(r, COL_TASK_KEY).Value = NextTaskKey()
            Dim colMap As HierarchyColumnMap
            colMap = TasksManagerColMap()
            ' Blank PIC on a new row -> appends at the end of the
            ' unassigned group under this (blank) parent, per
            ' NextSiblingOrderForPIC's own documented semantics.
            ws.Cells(r, COL_SIBLING_ORDER).Value = NextSiblingOrderForPIC(ws, colMap, "", "")

            ws.Cells(r, COL_UP).Value = ICON_UP
            ws.Cells(r, COL_DOWN).Value = ICON_DOWN
            ' Completed Subtasks stays blank (not 0) on a brand-new row --
            ' matches Total Subtasks' own "blank not zero" convention;
            ' Phase 5's CascadeComplete is what first writes a real value.

            RenumberHierarchy ws, colMap
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

' New Worksheet_Change case Priority doesn't need (plan §2): a PIC edit
' must re-sort its row to the end of its new PIC cluster and trigger a
' full renumber, not just sit as a plain value. Un-guarded by TaskKey
' (like HandleLinkedCampaignChange) -- RenumberHierarchy is already a
' no-op for a row with no TaskKey yet (modHierarchy.bas skips blank-Key
' rows in Pass 1), so running it here even on a not-yet-named row is
' harmless; the PIC value itself is preserved either way and takes effect
' once HandleTaskNameChange's own RenumberHierarchy call runs for real.
Public Function HandlePICChange(r As Long) As String
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")

    Dim colMap As HierarchyColumnMap
    colMap = TasksManagerColMap()

    Dim parentKey As String
    parentKey = Trim$(ws.Cells(r, COL_PARENT_KEY).Value & "")
    Dim newPIC As String
    newPIC = Trim$(ws.Cells(r, COL_PIC).Value & "")

    ' Clearing PIC back to blank moves the task to the end of the
    ' unassigned group (not back to its pre-assignment position) -- the
    ' natural minimal-surprise consequence of "always append at cluster
    ' end," same as how reparenting already works today (plan §4).
    ws.Cells(r, COL_SIBLING_ORDER).Value = NextSiblingOrderForPIC(ws, colMap, parentKey, newPIC)

    RenumberHierarchy ws, colMap
    RefreshButtonStyling ws

    Application.EnableEvents = prevEvents
    HandlePICChange = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandlePICChange = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Computes the next monotonic TaskKey, checking both Tasks (live) and
' Archive (frozen) -- see ARCHIVE_COL_TASK_KEY's own comment for why 16.
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
            RenumberHierarchy ws, TasksManagerColMap()   ' restores the displayed Subtask of text to its real current value
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
        colMap = TasksManagerColMap()
        ' Scoped by (newParentKey, this task's own current PIC) -- a
        ' reparent never changes PIC, so the task should land at the end
        ' of its own PIC cluster under its new parent (plan §4).
        Dim myPIC As String
        myPIC = Trim$(ws.Cells(r, COL_PIC).Value & "")
        ws.Cells(r, COL_SIBLING_ORDER).Value = NextSiblingOrderForPIC(ws, colMap, newParentKey, myPIC)
    End If

    RenumberHierarchy ws, TasksManagerColMap()
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

' Collects a task row and every descendant row (ID-prefix match).
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

' Archive's table starts at the "minimum valid size" (header + one blank
' data row) convention, same as employee's Archive -- reuses that first
' row directly the one time it's still genuinely blank. Falls back to
' plain ListRows.Add otherwise. Same pattern as modTasks.bas's
' NextArchiveRow (see its own comment); duplicated here rather than
' shared, consistent with this module being a genuine fork.
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

' Shared helper (plan §6) routed through by BOTH the real-completion path
' (ArchiveCompletedRow below) and the Parent-Context-snapshot path
' (ArchiveParentContextRow below) -- this single chokepoint is what makes
' "an existing Parent-Context row converts to Completed in place" fall out
' for free with no separate conversion logic: when an ancestor later
' completes for real, ArchiveCompletedRow's own call here naturally finds
' and overwrites whatever Parent-Context row an earlier sibling's
' completion already created for it, keyed purely by TaskKey (never by
' row position, which is meaningless across two separate completions).
' Application.Match (not WorksheetFunction.Match) deliberately -- it
' returns an error VARIANT on no match rather than raising, so a plain
' IsError check is enough, no error handling needed for the expected
' "not found yet" case.
Private Function FindOrCreateArchiveRow(archTbl As ListObject, taskKeyValue As Double) As ListRow
    If archTbl.ListRows.Count = 0 Then
        Set FindOrCreateArchiveRow = archTbl.ListRows.Add
        Exit Function
    End If

    Dim matchResult As Variant
    matchResult = Application.Match(taskKeyValue, archTbl.ListColumns("TaskKey").DataBodyRange, 0)

    If Not IsError(matchResult) Then
        Set FindOrCreateArchiveRow = archTbl.ListRows(CLng(matchResult))
    Else
        Set FindOrCreateArchiveRow = NextArchiveRow(archTbl)
    End If
End Function

' Writes/overwrites (via FindOrCreateArchiveRow) taskRow's own Archive row
' as RowType="Completed" -- the branch row that's actually completing (as
' opposed to ArchiveParentContextRow below, for a still-open ancestor).
' archivedCompletedSubtasks is precomputed by CascadeComplete's own §5-
' step-1 pass (own live counter + count of other branch rows descendant of
' this one), not read from the live sheet here.
Private Sub ArchiveCompletedRow(tasksWs As Worksheet, taskRow As Long, archTbl As ListObject, completedAt As Date, archivedCompletedSubtasks As Double, picLookup As Object)
    Dim taskKeyVal As Double
    taskKeyVal = CDbl(tasksWs.Cells(taskRow, COL_TASK_KEY).Value)

    Dim newRow As ListRow
    Set newRow = FindOrCreateArchiveRow(archTbl, taskKeyVal)

    newRow.Range(1, ARCHIVE_COL_ID).Value = tasksWs.Cells(taskRow, COL_ID).Value
    newRow.Range(1, ARCHIVE_COL_START_DATE).Value = tasksWs.Cells(taskRow, COL_START_DATE).Value
    newRow.Range(1, ARCHIVE_COL_TASK).Value = tasksWs.Cells(taskRow, COL_TASK).Value
    Dim picVal As String
    picVal = CStr(tasksWs.Cells(taskRow, COL_PIC).Value & "")
    newRow.Range(1, ARCHIVE_COL_PIC).Value = picVal
    ApplyPICColorToCell newRow.Range(1, ARCHIVE_COL_PIC), picVal, picLookup
    newRow.Range(1, ARCHIVE_COL_TOTAL_SUBTASKS).Value = tasksWs.Cells(taskRow, COL_NUM_SUBTASKS).Value
    newRow.Range(1, ARCHIVE_COL_COMPLETED_SUBTASKS).Value = archivedCompletedSubtasks
    newRow.Range(1, ARCHIVE_COL_SUBTASK_OF).Value = tasksWs.Cells(taskRow, COL_SUBTASK_OF).Value
    newRow.Range(1, ARCHIVE_COL_PROJECT).Value = tasksWs.Cells(taskRow, COL_PROJECT).Value
    newRow.Range(1, ARCHIVE_COL_LINKED_CAMPAIGN).Value = tasksWs.Cells(taskRow, COL_LINKED_CAMPAIGN).Value
    newRow.Range(1, ARCHIVE_COL_PRIORITY).Value = tasksWs.Cells(taskRow, COL_PRIORITY).Value
    newRow.Range(1, ARCHIVE_COL_DEADLINE).Value = tasksWs.Cells(taskRow, COL_DEADLINE).Value
    newRow.Range(1, ARCHIVE_COL_STATUS).Value = "Completed"
    ' Days Left = whatever Tasks' own Days Left formula currently reads,
    ' frozen -- plan §6 gives this semantic explicitly for the Parent-
    ' Context case ("same arithmetic Tasks itself uses"); the manager
    ' Archive layout only has one "Days Left" column (not employee's
    ' separate "days early/late vs completion" metric), so the Completed
    ' row uses the identical snapshot, for consistency between row types.
    newRow.Range(1, ARCHIVE_COL_DAYS_LEFT).Value = tasksWs.Cells(taskRow, COL_DAYS_LEFT).Value
    newRow.Range(1, ARCHIVE_COL_NOTES).Value = tasksWs.Cells(taskRow, COL_NOTES).Value
    newRow.Range(1, ARCHIVE_COL_COMPLETED_AT).Value = completedAt
    newRow.Range(1, ARCHIVE_COL_TASK_KEY).Value = taskKeyVal
    newRow.Range(1, ARCHIVE_COL_PARENT_KEY).Value = tasksWs.Cells(taskRow, COL_PARENT_KEY).Value
    newRow.Range(1, ARCHIVE_COL_LEVEL).Value = tasksWs.Cells(taskRow, COL_LEVEL).Value
    ' Visual tree structure (user-requested after using the manager
    ' build, alongside the append-order fix in CascadeComplete): mirrors
    ' Tasks' own RenumberHierarchy, which sets this same property on its
    ' Task cell so nesting reads visually, reinforced by archive.py's
    ' depth-shading/"⤷ " CF (both keyed off this row's own frozen Level).
    newRow.Range(1, ARCHIVE_COL_TASK).IndentLevel = tasksWs.Cells(taskRow, COL_LEVEL).Value
    newRow.Range(1, ARCHIVE_COL_ROW_TYPE).Value = "Completed"
End Sub

' Writes/updates (via FindOrCreateArchiveRow) a still-open ancestor's
' Parent-Context snapshot row -- plan §6's values: Status is the
' ancestor's actual CURRENT status (refreshed on every update, since the
' ancestor is still live in Tasks); Completed At stays blank (it hasn't
' completed); Total/Completed Subtasks are the ancestor's current live
' values (Completed Subtasks already reflects CascadeComplete's own
' increment by the time this is called -- see call site below).
Private Sub ArchiveParentContextRow(tasksWs As Worksheet, ancestorRow As Long, archTbl As ListObject, picLookup As Object)
    Dim taskKeyVal As Double
    taskKeyVal = CDbl(tasksWs.Cells(ancestorRow, COL_TASK_KEY).Value)

    Dim newRow As ListRow
    Set newRow = FindOrCreateArchiveRow(archTbl, taskKeyVal)

    newRow.Range(1, ARCHIVE_COL_ID).Value = tasksWs.Cells(ancestorRow, COL_ID).Value
    newRow.Range(1, ARCHIVE_COL_START_DATE).Value = tasksWs.Cells(ancestorRow, COL_START_DATE).Value
    newRow.Range(1, ARCHIVE_COL_TASK).Value = tasksWs.Cells(ancestorRow, COL_TASK).Value
    Dim ancestorPicVal As String
    ancestorPicVal = CStr(tasksWs.Cells(ancestorRow, COL_PIC).Value & "")
    newRow.Range(1, ARCHIVE_COL_PIC).Value = ancestorPicVal
    ApplyPICColorToCell newRow.Range(1, ARCHIVE_COL_PIC), ancestorPicVal, picLookup
    newRow.Range(1, ARCHIVE_COL_TOTAL_SUBTASKS).Value = tasksWs.Cells(ancestorRow, COL_NUM_SUBTASKS).Value
    newRow.Range(1, ARCHIVE_COL_COMPLETED_SUBTASKS).Value = tasksWs.Cells(ancestorRow, COL_COMPLETED_SUBTASKS).Value
    newRow.Range(1, ARCHIVE_COL_SUBTASK_OF).Value = tasksWs.Cells(ancestorRow, COL_SUBTASK_OF).Value
    newRow.Range(1, ARCHIVE_COL_PROJECT).Value = tasksWs.Cells(ancestorRow, COL_PROJECT).Value
    newRow.Range(1, ARCHIVE_COL_LINKED_CAMPAIGN).Value = tasksWs.Cells(ancestorRow, COL_LINKED_CAMPAIGN).Value
    newRow.Range(1, ARCHIVE_COL_PRIORITY).Value = tasksWs.Cells(ancestorRow, COL_PRIORITY).Value
    newRow.Range(1, ARCHIVE_COL_DEADLINE).Value = tasksWs.Cells(ancestorRow, COL_DEADLINE).Value
    newRow.Range(1, ARCHIVE_COL_STATUS).Value = tasksWs.Cells(ancestorRow, COL_STATUS).Value
    newRow.Range(1, ARCHIVE_COL_DAYS_LEFT).Value = tasksWs.Cells(ancestorRow, COL_DAYS_LEFT).Value
    newRow.Range(1, ARCHIVE_COL_NOTES).Value = tasksWs.Cells(ancestorRow, COL_NOTES).Value
    newRow.Range(1, ARCHIVE_COL_COMPLETED_AT).Value = ""
    newRow.Range(1, ARCHIVE_COL_TASK_KEY).Value = taskKeyVal
    newRow.Range(1, ARCHIVE_COL_PARENT_KEY).Value = tasksWs.Cells(ancestorRow, COL_PARENT_KEY).Value
    newRow.Range(1, ARCHIVE_COL_LEVEL).Value = tasksWs.Cells(ancestorRow, COL_LEVEL).Value
    newRow.Range(1, ARCHIVE_COL_TASK).IndentLevel = tasksWs.Cells(ancestorRow, COL_LEVEL).Value
    newRow.Range(1, ARCHIVE_COL_ROW_TYPE).Value = "Parent Context"
End Sub

' Restructured from modTasks.bas's CascadeComplete (same overall shape:
' EnableEvents guard -> archive loop -> delete bottom-up -> RenumberHierarchy
' -> RefreshButtonStyling), with plan §5's counter/ancestor-walk logic
' inserted between the archive loop and the delete loop -- must run BEFORE
' delete so ParentKey chains up from the branch are still resolvable.
' Private, not Public -- unlike CascadeDelete, HandleStatusChange's
' "Completed" case has no MsgBox confirmation to bypass for testing, so
' there's no testability reason to widen this (same as modTasks.bas's own
' CascadeComplete, also Private).
Private Sub CascadeComplete(ws As Worksheet, r As Long)
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim rowsInBranch() As Long
    Dim n As Long
    FindBranchRows ws, r, rowsInBranch, n

    ' §5 step 1: for every row in the branch, archived Completed-Subtasks =
    ' its own live counter + count of OTHER branch rows whose ID is a
    ' descendant of its own ID (same ID-prefix technique FindBranchRows
    ' itself uses). Keyed by row number in a Dictionary, computed BEFORE
    ' SortAscending reorders rowsInBranch below -- row NUMBERS don't change
    ' from sorting the array, only their positions within it, but keying by
    ' row number sidesteps that distinction entirely rather than relying on
    ' it.
    Dim archivedCountByRow As Object
    Set archivedCountByRow = CreateObject("Scripting.Dictionary")

    Dim branchIDs() As String
    ReDim branchIDs(1 To n)
    Dim i As Long
    For i = 1 To n
        branchIDs(i) = Trim$(ws.Cells(rowsInBranch(i), COL_ID).Value & "")
    Next i

    For i = 1 To n
        Dim ownLive As Double
        ownLive = 0
        If Len(Trim$(ws.Cells(rowsInBranch(i), COL_COMPLETED_SUBTASKS).Value & "")) > 0 Then
            ownLive = CDbl(ws.Cells(rowsInBranch(i), COL_COMPLETED_SUBTASKS).Value)
        End If

        Dim descendantCount As Long
        descendantCount = 0
        Dim j As Long
        For j = 1 To n
            If j <> i Then
                If Left$(branchIDs(j), Len(branchIDs(i)) + 1) = branchIDs(i) & "." Then
                    descendantCount = descendantCount + 1
                End If
            End If
        Next j

        archivedCountByRow(rowsInBranch(i)) = ownLive + descendantCount
    Next i

    ' Capture the branch's own top row's ParentKey BEFORE anything moves
    ' or gets deleted.
    Dim topParentKey As String
    topParentKey = Trim$(ws.Cells(r, COL_PARENT_KEY).Value & "")

    Dim completedAt As Date
    completedAt = Now()

    Dim archiveWs As Worksheet
    Set archiveWs = ThisWorkbook.Sheets("Archive")
    archiveWs.Unprotect

    Dim archTbl As ListObject
    Set archTbl = archiveWs.ListObjects("ArchiveTable")

    ' Built once per completion event, not per row -- passed into both
    ' Archive-row writers below so a newly-archived row's PIC cell picks
    ' up her current color assignment immediately, without waiting for a
    ' manual "Refresh PIC Colors" click.
    Dim picLookup As Object
    Set picLookup = PICColorLookup()

    ' §5 step 2: walk up from the branch's own top row's ParentKey to the
    ' root, incrementing each still-present ancestor's live Completed-
    ' Subtasks counter by n and COLLECTING its row number -- but NOT
    ' writing its Archive row yet. Dangling references stop the walk the
    ' same way WouldCreateCycle treats them -- guard caps it as a final
    ' backstop.
    '
    ' > **Corrected after live use:** originally each ancestor's Archive
    ' > row was written immediately during this same walk (immediate
    ' > parent first, root last), and the branch's own rows were archived
    ' > BEFORE this walk even started. Since Archive is append-only, that
    ' > meant the completing task's own row always landed ABOVE its
    ' > parent's context row -- backwards from the tree (parent should
    ' > read above child), and the user reported exactly this after using
    ' > the manager build. Fixed by collecting the ancestor chain first,
    ' > then writing it in REVERSE (root down to immediate parent) BEFORE
    ' > the branch's own rows -- so newly appended rows read top-down as
    ' > root -> ... -> immediate parent -> completed branch, matching the
    ' > tree's real shape. An ancestor whose Archive row already existed
    ' > (from an earlier, unrelated completion) still just updates in
    ' > place via FindOrCreateArchiveRow -- its ROW POSITION reflects
    ' > whenever it was first created, not this event, which is correct:
    ' > only newly-created rows need the ordering fix.
    Dim ancestorRows() As Long
    Dim ancestorCount As Long
    ancestorCount = 0
    ReDim ancestorRows(1 To 1000)

    Dim currentParentKeyText As String
    currentParentKeyText = topParentKey
    Dim guard As Long
    guard = 0
    Do While Len(currentParentKeyText) > 0 And guard < 10000
        Dim ancestorRow As Long
        ancestorRow = FindRowByTaskKey(ws, CDbl(currentParentKeyText))
        If ancestorRow = 0 Then Exit Do

        Dim ancestorLive As Double
        ancestorLive = 0
        If Len(Trim$(ws.Cells(ancestorRow, COL_COMPLETED_SUBTASKS).Value & "")) > 0 Then
            ancestorLive = CDbl(ws.Cells(ancestorRow, COL_COMPLETED_SUBTASKS).Value)
        End If
        ws.Cells(ancestorRow, COL_COMPLETED_SUBTASKS).Value = ancestorLive + n

        ancestorCount = ancestorCount + 1
        ancestorRows(ancestorCount) = ancestorRow

        currentParentKeyText = Trim$(ws.Cells(ancestorRow, COL_PARENT_KEY).Value & "")
        guard = guard + 1
    Loop

    ' Write ancestor context rows root-first (reverse of collection order),
    ' then the branch's own completed rows -- see correction note above.
    Dim m As Long
    For m = ancestorCount To 1 Step -1
        ArchiveParentContextRow ws, ancestorRows(m), archTbl, picLookup
    Next m

    SortAscending rowsInBranch, n   ' archive in visual (depth-first) order
    Dim k As Long
    For k = 1 To n
        Dim thisRow As Long
        thisRow = rowsInBranch(k)
        ArchiveCompletedRow ws, thisRow, archTbl, completedAt, CDbl(archivedCountByRow(thisRow)), picLookup
    Next k

    archiveWs.Protect UserInterfaceOnly:=True

    SortDescending rowsInBranch, n   ' delete bottom-up so row indexes stay valid
    For k = 1 To n
        ws.Rows(rowsInBranch(k)).Delete
    Next k

    RenumberHierarchy ws, TasksManagerColMap()
    RefreshButtonStyling ws
    RefreshLinkedCampaignLabels ws

    Application.EnableEvents = prevEvents
End Sub

' Public (not Private) so it's callable directly via Application.Run during
' testing, bypassing HandleStatusChange's MsgBox confirmation -- same
' convention as modTasks.bas's CascadeDelete/WouldCreateCycle.
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

    RenumberHierarchy ws, TasksManagerColMap()
    RefreshButtonStyling ws
    RefreshLinkedCampaignLabels ws

    Application.EnableEvents = prevEvents
    Exit Sub
ErrHandler:
    Application.EnableEvents = prevEvents
    Err.Raise Err.Number, , Err.Description
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

' Same neighbor-scan as modTasks.bas's HandleSiblingReorder, plus one added
' condition (plan §4): once the scan finds the row matching Level+
' ParentKey (the true physical neighbor), it must ALSO match PIC before
' being accepted as targetRow. If PIC differs, the loop still exits at
' that row (it genuinely is the physical neighbor -- scanning further
' would skip past it incorrectly) but targetRow stays 0, so the existing
' "NOOP" path below fires unchanged -- a PIC-cluster boundary behaves like
' already being first/last within the cluster.
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
    Dim myPIC As String
    myPIC = Trim$(ws.Cells(r, COL_PIC).Value & "")

    Dim targetRow As Long
    targetRow = 0

    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)

    If LCase$(direction) = "up" Then
        Dim i As Long
        For i = r - 1 To FIRST_DATA_ROW Step -1
            If CLng(ws.Cells(i, COL_LEVEL).Value) < myLevel Then Exit For
            If CLng(ws.Cells(i, COL_LEVEL).Value) = myLevel And Trim$(ws.Cells(i, COL_PARENT_KEY).Value & "") = myParentKey Then
                If StrComp(Trim$(ws.Cells(i, COL_PIC).Value & ""), myPIC, vbTextCompare) = 0 Then
                    targetRow = i
                End If
                Exit For
            End If
        Next i
    Else
        Dim j As Long
        For j = r + 1 To lastRow
            If CLng(ws.Cells(j, COL_LEVEL).Value) < myLevel Then Exit For
            If CLng(ws.Cells(j, COL_LEVEL).Value) = myLevel And Trim$(ws.Cells(j, COL_PARENT_KEY).Value & "") = myParentKey Then
                If StrComp(Trim$(ws.Cells(j, COL_PIC).Value & ""), myPIC, vbTextCompare) = 0 Then
                    targetRow = j
                End If
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

    RenumberHierarchy ws, TasksManagerColMap()
    RefreshButtonStyling ws
    RefreshLinkedCampaignLabels ws

    Application.EnableEvents = prevEvents
    HandleSiblingReorder = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandleSiblingReorder = "ERROR " & Err.Number & ": " & Err.Description
End Function

' ============================================================================
' Small helpers
' ============================================================================

' Also applies each row's PIC cell color (user-requested: "have each PIC
' with different color") -- folded into this same re-scan-and-reconcile
' pass rather than a separate full-sheet scan, since this Sub already
' runs at every site a structural change (new task, reparent, reorder,
' delete, PIC change) needs a full re-render anyway. Builds the color
' lookup ONCE per call, not per row.
Public Sub RefreshButtonStyling(ws As Worksheet)
    Dim lookup As Object
    Set lookup = PICColorLookup()

    Dim lastRow As Long, r As Long
    lastRow = LastHierarchyRow(ws, COL_TASK_KEY, FIRST_DATA_ROW)
    For r = FIRST_DATA_ROW To lastRow
        If Len(Trim$(ws.Cells(r, COL_TASK_KEY).Value & "")) > 0 Then
            modDesignSystem.ApplyButtonLook ws.Cells(r, COL_UP)
            modDesignSystem.ApplyButtonLook ws.Cells(r, COL_DOWN)
            ApplyPICColorToCell ws.Cells(r, COL_PIC), CStr(ws.Cells(r, COL_PIC).Value & ""), lookup
        End If
    Next r
End Sub

' Scans PIC List's Name/Color columns and returns a Dictionary of
' UCase(name) -> Long RGB color, for every name whose Color swatch cell
' actually has a fill set. A swatch left "No Fill" (Interior.Pattern =
' xlNone) is deliberately excluded, not stored as some default color --
' Interior.Color on an unfilled cell still reads back SOME number (often
' 0/black), which would otherwise misreport "no color chosen" as "colored
' black." Rebuilt fresh on every call (PIC List's own roster/colors are
' small, ~200 rows max, and this only runs at real structural-change
' points, never in a tight per-keystroke loop) rather than cached, so a
' color change is picked up the moment the next refresh runs.
Private Function PICColorLookup() As Object
    Dim lookup As Object
    Set lookup = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PIC List")

    Dim r As Long
    For r = PIC_FIRST_ROW To PIC_LAST_ROW
        Dim nameText As String
        nameText = Trim$(ws.Cells(r, PIC_COL_NAME).Value & "")
        If Len(nameText) > 0 Then
            If ws.Cells(r, PIC_COL_COLOR).Interior.Pattern <> xlNone Then
                lookup(UCase$(nameText)) = ws.Cells(r, PIC_COL_COLOR).Interior.Color
            End If
        End If
    Next r

    Set PICColorLookup = lookup
End Function

' Sets targetCell's fill to picName's assigned color (font switched to
' black/white for legibility against whatever color she picked, see
' TextColorForBackground), or clears any existing fill/font-color override
' if picName has no assigned color (blank PIC, or a name whose swatch is
' "No Fill") -- explicit clear, not "leave whatever was there," so a task
' reassigned away from a colored PIC never keeps a stale color.
Private Sub ApplyPICColorToCell(targetCell As Range, picName As String, lookup As Object)
    Dim key As String
    key = UCase$(Trim$(picName))

    If Len(key) > 0 And lookup.Exists(key) Then
        Dim bgColor As Long
        bgColor = CLng(lookup(key))
        targetCell.Interior.Pattern = xlSolid
        targetCell.Interior.Color = bgColor
        targetCell.Font.Color = TextColorForBackground(bgColor)
    Else
        targetCell.Interior.Pattern = xlNone
        targetCell.Font.Color = modDesignSystem.COLOR_INK_PRIMARY
    End If
End Sub

' Simple perceived-brightness check (standard broadcast-luma weights) --
' picks black or white text so it stays readable against whatever
' arbitrary color she chooses, from pale yellow to dark navy.
Private Function TextColorForBackground(bgColor As Long) As Long
    Dim red As Long, green As Long, blue As Long
    red = bgColor Mod 256
    green = (bgColor \ 256) Mod 256
    blue = (bgColor \ 65536) Mod 256

    Dim luma As Double
    luma = (0.299 * red + 0.587 * green + 0.114 * blue) / 255

    If luma > 0.6 Then
        TextColorForBackground = RGB(0, 0, 0)
    Else
        TextColorForBackground = RGB(255, 255, 255)
    End If
End Function

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
    Dim i As Long, j As Long, tmp As Long
    For i = 2 To n
        tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j) < tmp Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = tmp
    Next i
End Sub

' ============================================================================
' PIC List recompaction + reorder (deferred from Phase 2 -- see
' sheets/pic_list.py's own docstring; Up/Down added after live use of the
' manager build). RecompactPICList closes a gap left by clearing a name
' mid-list, so PICList's COUNTA-driven dynamic range never undercounts
' past a gap (confirmed live in Phase 2 that a cleared middle name
' otherwise drops every entry after it out of the dropdown). Called from
' PIC List.bas's Worksheet_Change on every edit to column A; also called
' defensively from Workbook_Open (ThisWorkbookManager.bas), alongside
' RefreshPICListButtonStyling.
' ============================================================================

Public Function RecompactPICList() As String
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PIC List")

    Dim vals As Variant
    vals = ws.Range(ws.Cells(PIC_FIRST_ROW, PIC_COL_NAME), ws.Cells(PIC_LAST_ROW, PIC_COL_NAME)).Value

    Dim outVals() As Variant
    ReDim outVals(1 To (PIC_LAST_ROW - PIC_FIRST_ROW + 1), 1 To 1)

    Dim i As Long, outIdx As Long
    outIdx = 0
    For i = 1 To UBound(vals, 1)
        Dim nameText As String
        nameText = Trim$(vals(i, 1) & "")
        If Len(nameText) > 0 Then
            outIdx = outIdx + 1
            outVals(outIdx, 1) = nameText
        End If
    Next i
    ' Remaining slots (outIdx+1 .. end) stay blank via ReDim's default --
    ' this is what actually closes the gap, since the write below
    ' overwrites the whole range unconditionally.

    ws.Range(ws.Cells(PIC_FIRST_ROW, PIC_COL_NAME), ws.Cells(PIC_LAST_ROW, PIC_COL_NAME)).Value = outVals

    Application.EnableEvents = prevEvents
    RecompactPICList = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    RecompactPICList = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Plain adjacent-row swap -- PIC List has no hierarchy/grouping (unlike
' Tasks' PIC-cluster-scoped HandleSiblingReorder), so "up"/"down" just
' means the physically adjacent row, full stop. A blank row (no name) is
' never a valid target, and a row at either end of the list correctly
' no-ops the same way Tasks' own reorder already does at a boundary.
' Only the Name column moves -- Up/Down are fixed per-row button glyphs,
' not per-name data, so they never need to travel with a swap.
Public Function HandlePICListReorder(r As Long, direction As String) As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PIC List")

    If Len(Trim$(ws.Cells(r, PIC_COL_NAME).Value & "")) = 0 Then
        HandlePICListReorder = "NOOP"
        Exit Function
    End If

    Dim targetRow As Long
    targetRow = 0
    If LCase$(direction) = "up" Then
        If r - 1 >= PIC_FIRST_ROW Then
            If Len(Trim$(ws.Cells(r - 1, PIC_COL_NAME).Value & "")) > 0 Then targetRow = r - 1
        End If
    Else
        If r + 1 <= PIC_LAST_ROW Then
            If Len(Trim$(ws.Cells(r + 1, PIC_COL_NAME).Value & "")) > 0 Then targetRow = r + 1
        End If
    End If

    If targetRow = 0 Then
        HandlePICListReorder = "NOOP"
        Exit Function
    End If

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim myName As String, targetName As String
    myName = ws.Cells(r, PIC_COL_NAME).Value
    targetName = ws.Cells(targetRow, PIC_COL_NAME).Value
    ws.Cells(r, PIC_COL_NAME).Value = targetName
    ws.Cells(targetRow, PIC_COL_NAME).Value = myName

    Application.EnableEvents = prevEvents
    HandlePICListReorder = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandlePICListReorder = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Re-scan-and-reconcile, same convention as modTasksManager.RefreshButtonStyling
' for Tasks itself: every row with a real name gets the Up/Down button
' look (and its icon glyphs, in case a row just became occupied for the
' first time); every row without one gets fully un-styled via
' ClearButtonLook, so a name cleared (directly, or shifted out by
' RecompactPICList) never leaves stale button borders behind on what's
' now an empty row -- PIC List has no physical Rows.Delete the way
' Tasks'/Recurring Tasks' hierarchy engine does, so a vacated row would
' otherwise keep its old formatting forever (same reasoning as Campaigns'
' CompactCampaignSlots, see modDesignSystem.ClearButtonLook's own
' comment).
Public Function RefreshPICListButtonStyling() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PIC List")

    Dim r As Long
    For r = PIC_FIRST_ROW To PIC_LAST_ROW
        If Len(Trim$(ws.Cells(r, PIC_COL_NAME).Value & "")) > 0 Then
            ws.Cells(r, PIC_COL_UP).Value = ICON_UP
            ws.Cells(r, PIC_COL_DOWN).Value = ICON_DOWN
            modDesignSystem.ApplyButtonLook ws.Cells(r, PIC_COL_UP)
            modDesignSystem.ApplyButtonLook ws.Cells(r, PIC_COL_DOWN)
        Else
            ws.Cells(r, PIC_COL_UP).Value = ""
            ws.Cells(r, PIC_COL_DOWN).Value = ""
            modDesignSystem.ClearButtonLook ws.Cells(r, PIC_COL_UP)
            modDesignSystem.ClearButtonLook ws.Cells(r, PIC_COL_DOWN)
        End If
    Next r

    RefreshPICListButtonStyling = "OK"
    Exit Function
ErrHandler:
    RefreshPICListButtonStyling = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Button handler for PIC List's "Refresh PIC Colors" banner (plan/user
' feedback: "a way for her to change the color of each PIC"). Necessary
' because Excel/VBA has no event that fires on a FORMATTING-only change --
' Worksheet_Change never sees her picking a Fill Color on a swatch cell --
' so a color change can't propagate automatically the instant she makes
' it; this is the explicit, immediate alternative (ThisWorkbookManager.bas's
' Workbook_Open also runs it once as a backfill, same convention as
' RecompactPICList/RefreshPICListButtonStyling there).
'
' Tasks: reuses RefreshButtonStyling, which already folds PIC-color
' application into its own re-scan-and-reconcile pass -- no separate scan
' needed. Archive: a dedicated re-scan of every EXISTING archived row (not
' just rows newly written since the last click), since CascadeComplete's
' own automatic coloring only ever touches a row at the moment it's
' archived -- a color changed afterward needs this explicit refresh to
' reach already-archived rows.
Public Function HandleRefreshPICColorsClick() As String
    On Error GoTo ErrHandler
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    RefreshButtonStyling ThisWorkbook.Sheets("Tasks")

    Dim lookup As Object
    Set lookup = PICColorLookup()

    Dim archiveWs As Worksheet
    Set archiveWs = ThisWorkbook.Sheets("Archive")
    archiveWs.Unprotect

    Dim lastRow As Long, r As Long
    lastRow = LastHierarchyRow(archiveWs, ARCHIVE_COL_TASK_KEY, 2)
    For r = 2 To lastRow
        If Len(Trim$(archiveWs.Cells(r, ARCHIVE_COL_TASK).Value & "")) > 0 Then
            Dim picVal As String
            picVal = CStr(archiveWs.Cells(r, ARCHIVE_COL_PIC).Value & "")
            ApplyPICColorToCell archiveWs.Cells(r, ARCHIVE_COL_PIC), picVal, lookup
        End If
    Next r

    archiveWs.Protect UserInterfaceOnly:=True

    Application.EnableEvents = prevEvents
    HandleRefreshPICColorsClick = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandleRefreshPICColorsClick = "ERROR " & Err.Number & ": " & Err.Description
End Function

' ============================================================================
' Row extender -- adapted copy of modTasks.bas's own (see its header
' comment for the full design rationale; not repeated here). Differences:
' no Timer/Time Spent formatting or formulas; re-applies the PIC dropdown
' to new rows; Completed Subtasks left blank on new rows (matches Total
' Subtasks' "blank not zero" convention); every ScreenUpdating/Calculation
' guard included from day one (plan §4), not retrofitted like modTasks.bas
' originally needed.
' ============================================================================

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

Public Function RowExtenderBannerRow() As Long
    RowExtenderBannerRow = TasksRowAllowance() + 1
End Function

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

    Dim bannerRange As Range
    Set bannerRange = ws.Range(ws.Cells(oldLast + 1, 1), ws.Cells(oldLast + 1, 6))
    bannerRange.UnMerge
    bannerRange.ClearFormats
    bannerRange.ClearContents

    Dim r As Long
    For r = oldLast + 1 To newLast
        BuildTaskTemplateRow ws, r, newLast
    Next r

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

' newLast is the CampaignRank/Total-Subtasks bound this row's formulas
' should use. Mirrors sheets/tasks.py's build() loop for variant="manager".
Private Sub BuildTaskTemplateRow(ws As Worksheet, r As Long, newLast As Long)
    Dim DQ As String
    DQ = Chr(34)

    ws.Cells(r, COL_START_DATE).NumberFormat = "DD.MM.YYYY HH:MM"
    ws.Cells(r, COL_DEADLINE).NumberFormat = "DD.MM.YYYY HH:MM"
    ws.Cells(r, COL_NOTES).NumberFormat = "General;General;General;" & DQ & modDesignSystem.ICON_NOTES() & DQ

    Dim idRange As String, thisId As String, countExpr As String
    idRange = "$A$2:$A$" & newLast
    thisId = "A" & r
    countExpr = "COUNTIF(" & idRange & "," & thisId & "&" & DQ & ".*" & DQ & ")"
    ws.Cells(r, COL_NUM_SUBTASKS).Formula = "=IF(" & countExpr & "=0," & DQ & DQ & "," & countExpr & ")"

    Dim dl As String
    dl = "M" & r   ' Deadline
    ws.Cells(r, COL_DAYS_LEFT).Formula = "=IF(" & dl & "=" & DQ & DQ & "," & DQ & DQ & ",INT(" & dl & ")-TODAY())"

    Dim key As String
    key = "Q" & r   ' TaskKey
    ws.Cells(r, COL_SORT_KEY).Formula = "=IF(" & dl & "=" & DQ & DQ & "," & DQ & DQ & "," & dl & "+" & key & "/1000000)"

    Dim lck As String
    lck = "U" & r   ' LinkedCampaignKey
    ws.Cells(r, COL_CAMPAIGN_SORT_HELPER).Formula = _
        "=IF(" & lck & "=" & DQ & DQ & "," & DQ & DQ & ",IF(" & dl & "=" & DQ & DQ & ",9999999+" & key & "/1000000," & _
        dl & "+" & key & "/1000000))"

    Dim cshRange As String, lckRange As String, csh As String
    cshRange = "$W$2:$W$" & newLast
    lckRange = "$U$2:$U$" & newLast
    csh = "W" & r
    ws.Cells(r, COL_CAMPAIGN_RANK).Formula = _
        "=IF(" & lck & "=" & DQ & DQ & "," & DQ & DQ & ",SUMPRODUCT((" & lckRange & "=" & lck & ")*(" & _
        cshRange & "<=" & csh & ")))"

    Dim rankCell As String
    rankCell = "X" & r
    ws.Cells(r, COL_CAMPAIGN_RANK_KEY).Formula = _
        "=IF(" & lck & "=" & DQ & DQ & "," & DQ & DQ & "," & lck & "&" & DQ & "|" & DQ & "&" & rankCell & ")"

    Dim c As Variant
    For Each c In Array(COL_ID, COL_NUM_SUBTASKS, COL_COMPLETED_SUBTASKS, COL_SUBTASK_OF, COL_DAYS_LEFT)
        With ws.Cells(r, CLng(c))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_NOTES)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = modDesignSystem.COLOR_GRIDLINE
    End With
End Sub

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

Private Sub RewriteCampaignRankFormulas(ws As Worksheet, newLast As Long)
    Dim DQ As String
    DQ = Chr(34)
    Dim cshRange As String, lckRange As String
    cshRange = "$W$2:$W$" & newLast
    lckRange = "$U$2:$U$" & newLast
    Dim r As Long, lck As String, csh As String
    For r = FIRST_DATA_ROW To newLast
        lck = "U" & r
        csh = "W" & r
        ws.Cells(r, COL_CAMPAIGN_RANK).Formula = _
            "=IF(" & lck & "=" & DQ & DQ & "," & DQ & DQ & ",SUMPRODUCT((" & lckRange & "=" & lck & ")*(" & _
            cshRange & "<=" & csh & ")))"
    Next r
End Sub

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

    Set newRange = ws.Range(ws.Cells(oldLast + 1, COL_PIC), ws.Cells(newLast, COL_PIC))
    With newRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=PICList"
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

Private Sub ExtendCFAppliesTo(ws As Worksheet, firstCol As Long, lastCol As Long, newLast As Long)
    Dim newRange As Range
    Set newRange = ws.Range(ws.Cells(FIRST_DATA_ROW, firstCol), ws.Cells(newLast, lastCol))

    Dim srcConditions As FormatConditions
    Set srcConditions = ws.Cells(FIRST_DATA_ROW, firstCol).FormatConditions
    Dim n As Long, i As Long
    n = srcConditions.Count
    For i = 1 To n
        srcConditions.Item(i).ModifyAppliesToRange newRange
    Next i
End Sub

Private Sub ReplaceRowBound(ws As Worksheet, oldLast As Long, newLast As Long)
    ws.Cells.Replace What:="$" & oldLast, Replacement:="$" & newLast, LookAt:=xlPart, MatchCase:=True
End Sub

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
