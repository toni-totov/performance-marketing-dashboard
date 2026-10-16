Attribute VB_Name = "modMigration"
Option Explicit

' One-time, user-run tool (PLAN.md §21): copies real data out of an OLDER
' copy of this workbook (picked via a file dialog) into THIS workbook,
' matching columns by their HEADER TEXT rather than fixed position --
' resilient to schema drift between the old file and a freshly rebuilt
' one, which is exactly the situation this exists for (old file predates
' PLAN.md §20's fixes). Everything happens locally in the Excel process
' the user already has open: the old file is opened read-only and closed
' without saving, nothing is written back to it, nothing leaves this
' machine. Entry point: MigrateFromPrevious (Alt+F8; the old name
' MigrateFromV2 still works).
'
' Design: hierarchy sheets (Tasks, Recurring Tasks) only need their
' TRULY position-independent identity fields copied verbatim (stable key,
' parent key, sibling order) -- ID/Level/Subtask-of display text are all
' recomputed fresh by modHierarchy.RenumberHierarchy from those three
' alone (confirmed by reading RenumberHierarchy itself: it overwrites
' ColID/ColLevel/ColSubtaskOf unconditionally from the freshly-walked
' tree, never reads them as input), so this deliberately does NOT migrate
' them -- copying stale ID/Level text would just get thrown away anyway.
' Flat Table sheets (Archive, Campaigns Archive, Login Information, Quick
' Links, Schedule Data's two tables) reuse the project's own established
' "minimum valid size" convention (NextTableRow below) so the first
' pre-built blank row gets filled in rather than left permanently empty.
' Campaigns is neither -- a fixed 20-slot layout (modCampaigns.bas's own
' header comment) -- so its main rows are copied slot-by-slot and then
' modCampaigns.CompactCampaignSlots is reused to apply real button
' styling and pack any gaps, exactly as if the user had typed each
' campaign in by hand.

Private Const HEADER_SCAN_MAX_COL As Long = 60

Public gMigrationLog As String
Public gMigrationRow As Long
Public gMigrationHadError As Boolean

' Test-only accessor -- Application.Run can't read a Public module
' variable directly, only call a Function/Sub.
Public Function GetMigrationLogForTest() As String
    GetMigrationLogForTest = gMigrationLog
End Function

' Test-only: exercises the EXACT same sequence as MigrateFromV2 (same
' Workbooks.Open args, same order, same EnableEvents/ScreenUpdating/
' DisplayAlerts handling) but takes the path directly instead of showing
' the two dialogs, so a COM test can drive the real production code path
' end-to-end -- not just the individual Migrate* subs in isolation.
Public Function MigrateFromV2Test(srcPath As String) As String
    On Error GoTo ErrHandler
    gMigrationLog = ""

    Dim prevEvents As Boolean, prevScreen As Boolean, prevAlerts As Boolean
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(srcPath, ReadOnly:=True, UpdateLinks:=0)

    RunAllStages srcWb

    srcWb.Close SaveChanges:=False

    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Application.DisplayAlerts = prevAlerts
    MigrateFromV2Test = IIf(gMigrationHadError, "STAGE ERRORS: ", "OK: ") & gMigrationLog
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Application.DisplayAlerts = prevAlerts
    MigrateFromV2Test = "ERROR " & Err.Number & ": " & Err.Description & " | Progress so far: " & gMigrationLog
End Function

' Old entry-point name, kept so the Alt+F8 instructions the user already
' has still work.
Public Sub MigrateFromV2()
    MigrateFromPrevious
End Sub

Public Sub MigrateFromPrevious()
    On Error GoTo ErrHandler

    Dim prevEvents As Boolean, prevScreen As Boolean, prevAlerts As Boolean
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts

    gMigrationLog = ""

    Dim resp As VbMsgBoxResult
    resp = MsgBox("This copies your tasks, campaigns, recurring templates, schedule data, " & _
        "logins, and quick links out of an older workbook into THIS one." & vbCrLf & vbCrLf & _
        "Nothing is sent anywhere -- it all happens on this computer. The old file is opened " & _
        "read-only and is never modified or saved." & vbCrLf & vbCrLf & _
        "You'll be asked to pick that old file next. Continue?", vbYesNo + vbQuestion, "Migrate from previous workbook")
    If resp <> vbYes Then Exit Sub

    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename("Excel Workbooks (*.xlsm; *.xlsx),*.xlsm;*.xlsx", , "Select your OLD workbook to migrate from")
    If VarType(srcPath) = vbBoolean Then Exit Sub   ' user cancelled

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(CStr(srcPath), ReadOnly:=True, UpdateLinks:=0)

    RunAllStages srcWb

    srcWb.Close SaveChanges:=False

    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Application.DisplayAlerts = prevAlerts

    If gMigrationHadError Then
        MsgBox "Migration finished, but some parts hit an error (marked !! below). Everything else was copied." & _
            vbCrLf & vbCrLf & gMigrationLog & vbCrLf & _
            "Please send the !! line(s) -- they contain only the step, a row number and the error text, no data.", _
            vbExclamation, "Migration finished with errors"
        Exit Sub
    End If
    MsgBox "Migration finished." & vbCrLf & vbCrLf & gMigrationLog & vbCrLf & _
        "Review every sheet before you trust it -- then save this file yourself when you're happy.", _
        vbInformation, "Migration finished"
    Exit Sub

ErrHandler:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Application.DisplayAlerts = prevAlerts
    MsgBox "Migration stopped by an error before finishing:" & vbCrLf & Err.Number & ": " & Err.Description & _
        vbCrLf & vbCrLf & "Progress so far:" & vbCrLf & gMigrationLog & vbCrLf & _
        "Whatever copied before the error is already sitting in this open workbook. Close this file " & _
        "WITHOUT saving and start again from a fresh copy of the new template, rather than re-running " & _
        "on top of a half-migrated one.", vbCritical, "Migration error"
End Sub

' Scans a fixed, generous column bound rather than .End(xlToLeft) -- a
' hidden header column (TaskKey, ParentKey, etc.) is exactly the case
' CLAUDE.md documents .End() silently skipping past.
Private Function FindColumnByHeaderText(ws As Worksheet, headerRow As Long, label As String) As Long
    Dim c As Long
    For c = 1 To HEADER_SCAN_MAX_COL
        If Trim$(ws.Cells(headerRow, c).Value & "") = label Then
            FindColumnByHeaderText = c
            Exit Function
        End If
    Next c
    FindColumnByHeaderText = 0
End Function

' Same "reuse the untouched pre-built first row instead of always
' ListRows.Add-ing a new one" convention already used by modTasks.bas's
' NextArchiveRow and modCampaigns.bas's NextCampaignArchiveRow -- kept as
' its own small copy here rather than made public in either of those,
' since this is the only other caller and it's a 4-line check.
Private Function NextTableRow(tbl As ListObject) As ListRow
    If tbl.ListRows.Count = 1 Then
        If Len(Trim$(tbl.ListRows(1).Range(1, 1).Value & "")) = 0 Then
            Set NextTableRow = tbl.ListRows(1)
            Exit Function
        End If
    End If
    Set NextTableRow = tbl.ListRows.Add
End Function

' Destination column for a label, allowing for headers renamed since the
' old file was built (V4, PLAN.md §22.9).
Private Function FindDestColumn(ws As Worksheet, label As String) As Long
    FindDestColumn = FindColumnByHeaderText(ws, 1, label)
    If FindDestColumn = 0 And label = "Days Left" Then FindDestColumn = FindColumnByHeaderText(ws, 1, "Early / Late")
End Function

' Columns that only exist from V4 on -- an older source file can't have
' them, so their absence isn't worth a "missing" warning.
Private Function IsNewInV4(label As String) As Boolean
    Select Case label
        Case "Time", "If Missed", "RecurringSourceKey", "DoneSubtasks"
            IsNewInV4 = True
    End Select
End Function

Private Sub LogStageError(stage As String)
    gMigrationHadError = True
    gMigrationLog = gMigrationLog & "!! " & stage & " STOPPED"
    If gMigrationRow > 0 Then gMigrationLog = gMigrationLog & " at old-file row " & gMigrationRow
    gMigrationLog = gMigrationLog & ": error " & Err.Number & " -- " & Err.Description & vbCrLf
End Sub

' Every stage runs under its own error trap (V4, PLAN.md §22.9): one
' failing stage is reported -- with the stage, the old file's row and the
' exact error -- and the rest still run, instead of the whole migration
' stopping with a bare error number. Archive and Campaigns Archive are
' protected sheets, and Excel refuses Table operations (ListRows.Add) on a
' protected sheet even with UserInterfaceOnly -- PLAN.md §7.1. Before V4
' the migration never unprotected them, so the SECOND archived row (the
' first reuses the table's blank starter row) always failed with 1004
' "table functions aren't available because the sheet is protected" --
' the root cause of §21's open "MigrateFromV2 errored on the real file"
' bug, reproduced against a realistic V3 fixture.
Private Sub RunAllStages(srcWb As Workbook)
    Dim aws As Worksheet, caws As Worksheet
    Set aws = ThisWorkbook.Sheets("Archive")
    Set caws = ThisWorkbook.Sheets("Campaigns Archive")
    gMigrationHadError = False

    On Error Resume Next
    aws.Unprotect
    caws.Unprotect

    gMigrationRow = 0: Err.Clear
    MigrateTasksAndArchive srcWb
    If Err.Number <> 0 Then LogStageError "Tasks"
    gMigrationRow = 0: Err.Clear
    MigrateArchive srcWb
    If Err.Number <> 0 Then LogStageError "Archive"
    ' Subtask progress for migrated tasks, rebuilt from the migrated
    ' Archive (§22.11) -- needs both Tasks and Archive in place first.
    gMigrationRow = 0: Err.Clear
    Dim seeded As Long
    seeded = modTasks.SeedDoneSubtaskCounters()
    If Err.Number <> 0 Then
        LogStageError "Subtask progress"
    ElseIf seeded > 0 Then
        gMigrationLog = gMigrationLog & "Subtask progress: rebuilt from Archive for " & seeded & " task(s)." & vbCrLf
    End If
    gMigrationRow = 0: Err.Clear
    MigrateCampaignsAndArchive srcWb
    If Err.Number <> 0 Then LogStageError "Campaigns"
    gMigrationRow = 0: Err.Clear
    MigrateCampaignsArchive srcWb
    If Err.Number <> 0 Then LogStageError "Campaigns Archive"
    gMigrationRow = 0: Err.Clear
    MigrateRecurringTasks srcWb
    If Err.Number <> 0 Then LogStageError "Recurring Tasks"
    gMigrationRow = 0: Err.Clear
    MigrateFlatTable srcWb, "Login Information", "LoginInformationTable", _
        Array("Platform", "Website", "Username / Email", "Password", "Last Password Change", "Access", "Notes")
    If Err.Number <> 0 Then LogStageError "Login Information"
    gMigrationRow = 0: Err.Clear
    MigrateFlatTable srcWb, "Quick Links", "QuickLinksTable", _
        Array("Category / Group", "Link / URL", "Description", "Notes")
    If Err.Number <> 0 Then LogStageError "Quick Links"
    gMigrationRow = 0: Err.Clear
    MigrateFlatTable srcWb, "Schedule Data", "WorkLocationTable", Array("Date", "Location")
    If Err.Number <> 0 Then LogStageError "Schedule Data (work location)"
    gMigrationRow = 0: Err.Clear
    MigrateFlatTable srcWb, "Schedule Data", "MeetingsTable", _
        Array("Meeting Name", "Start Time", "End Time", "Recurrence", "Anchor Date", "Status", "MeetingKey")
    If Err.Number <> 0 Then LogStageError "Schedule Data (meetings)"

    Err.Clear
    aws.Protect UserInterfaceOnly:=True
    caws.Protect UserInterfaceOnly:=True
    modSchedule.RenderScheduleBadges
    modDailyLog.GoThisWeek
    modToday.RenderTodayPanel
    On Error GoTo 0
End Sub

Private Function TrySheet(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set TrySheet = wb.Sheets(sheetName)
    On Error GoTo 0
End Function

' Public (not Private) purely so each stage is independently callable via
' Application.Run for testing -- same "MsgBox-bypass" reasoning CLAUDE.md
' documents for other modules' core logic, since MigrateFromV2 itself
' pops a file-picker dialog and a confirmation MsgBox, neither drivable
' over COM.
Public Sub MigrateTasksAndArchive(srcWb As Workbook)
    Dim srcWs As Worksheet
    Set srcWs = TrySheet(srcWb, "Tasks")
    If srcWs Is Nothing Then
        gMigrationLog = gMigrationLog & "Tasks: SKIPPED (no 'Tasks' sheet found in the old file)." & vbCrLf
        Exit Sub
    End If
    Dim destWs As Worksheet
    Set destWs = ThisWorkbook.Sheets("Tasks")

    Dim keyCol As Long
    keyCol = FindColumnByHeaderText(srcWs, 1, "TaskKey")
    If keyCol = 0 Then
        gMigrationLog = gMigrationLog & "Tasks: SKIPPED (couldn't find the TaskKey column -- old file may be a much older/different format than expected)." & vbCrLf
        Exit Sub
    End If

    Dim srcLastRow As Long
    srcLastRow = modHierarchy.LastHierarchyRow(srcWs, keyCol, 2)

    Dim srcRowCount As Long, r As Long
    For r = 2 To srcLastRow
        gMigrationRow = r
        If Len(Trim$(srcWs.Cells(r, keyCol).Value & "")) > 0 Then srcRowCount = srcRowCount + 1
    Next r

    If srcRowCount = 0 Then
        gMigrationLog = gMigrationLog & "Tasks: nothing to migrate (0 real rows found)." & vbCrLf
        Exit Sub
    End If

    ' Grow this workbook's Tasks allowance using the already-tested row
    ' extender (PLAN.md §20 Bug #10) rather than writing new capacity
    ' logic -- extend until there's comfortable headroom left over, not
    ' just exactly enough.
    Dim guard As Long
    guard = 0
    Do While CLng(Application.Evaluate("TasksRowAllowance")) < srcRowCount + 10
        modTasks.ExtendTaskRows
        guard = guard + 1
        If guard > 50 Then Err.Raise vbObjectError + 1, , "Tasks row extender didn't reach the needed capacity after 50 runs -- aborting rather than looping forever."
    Loop

    Dim labels() As Variant
    labels = Array("Start Date", "Task", "Project / Category", "Priority", "Deadline", "Status", _
                    "Time Spent (hrs)", "Notes", "TaskKey", "ParentKey", "SiblingOrder", _
                    "LinkedCampaignKey", "TimerRunningSince", "RecurringSourceKey", "DoneSubtasks")

    Dim n As Long
    n = UBound(labels) - LBound(labels) + 1
    Dim srcCols() As Long, destCols() As Long
    ReDim srcCols(1 To n)
    ReDim destCols(1 To n)
    Dim i As Long, missing As String
    For i = 1 To n
        srcCols(i) = FindColumnByHeaderText(srcWs, 1, CStr(labels(LBound(labels) + i - 1)))
        destCols(i) = FindDestColumn(destWs, CStr(labels(LBound(labels) + i - 1)))
        If srcCols(i) = 0 And Not IsNewInV4(CStr(labels(LBound(labels) + i - 1))) Then missing = missing & CStr(labels(LBound(labels) + i - 1)) & ", "
    Next i

    Dim destRow As Long
    destRow = 2
    For r = 2 To srcLastRow
        gMigrationRow = r
        If Len(Trim$(srcWs.Cells(r, keyCol).Value & "")) > 0 Then
            For i = 1 To n
                If srcCols(i) > 0 And destCols(i) > 0 Then
                    destWs.Cells(destRow, destCols(i)).Value = srcWs.Cells(r, srcCols(i)).Value
                End If
            Next i
            destRow = destRow + 1
        End If
    Next r

    modHierarchy.RenumberHierarchy destWs, modTasks.TasksColMap()
    modTasks.RefreshButtonStyling destWs
    modTasks.RefreshLinkedCampaignLabels destWs

    gMigrationLog = gMigrationLog & "Tasks: migrated " & srcRowCount & " row(s)."
    If Len(missing) > 0 Then gMigrationLog = gMigrationLog & " (old file was missing: " & Left$(missing, Len(missing) - 2) & ")"
    gMigrationLog = gMigrationLog & vbCrLf
End Sub

Public Sub MigrateArchive(srcWb As Workbook)
    Dim srcWs As Worksheet
    Set srcWs = TrySheet(srcWb, "Archive")
    If srcWs Is Nothing Then
        gMigrationLog = gMigrationLog & "Archive: SKIPPED (sheet not found)." & vbCrLf
        Exit Sub
    End If
    Dim destWs As Worksheet
    Set destWs = ThisWorkbook.Sheets("Archive")

    Dim labels() As Variant
    labels = Array("ID", "Start Date", "Task", "Total Subtasks", "Subtask of", "Project / Category", _
                    "Linked Campaign", "Priority", "Deadline", "Status", "Days Left", "Time Spent (hrs)", _
                    "Notes", "Completed At", "TaskKey", "ParentKey", "Level")

    Dim anchorCol As Long
    anchorCol = FindColumnByHeaderText(srcWs, 1, "Task")
    If anchorCol = 0 Then
        gMigrationLog = gMigrationLog & "Archive: SKIPPED (couldn't find the 'Task' column)." & vbCrLf
        Exit Sub
    End If

    Dim srcLastRow As Long
    srcLastRow = modHierarchy.LastHierarchyRow(srcWs, anchorCol, 2)

    Dim destTbl As ListObject
    Set destTbl = destWs.ListObjects("ArchiveTable")

    Dim n As Long
    n = UBound(labels) - LBound(labels) + 1
    Dim srcCols() As Long, destCols() As Long
    ReDim srcCols(1 To n)
    ReDim destCols(1 To n)
    Dim i As Long, missing As String
    For i = 1 To n
        srcCols(i) = FindColumnByHeaderText(srcWs, 1, CStr(labels(LBound(labels) + i - 1)))
        destCols(i) = FindDestColumn(destWs, CStr(labels(LBound(labels) + i - 1)))
        If srcCols(i) = 0 And Not IsNewInV4(CStr(labels(LBound(labels) + i - 1))) Then missing = missing & CStr(labels(LBound(labels) + i - 1)) & ", "
    Next i

    Dim r As Long, migrated As Long
    For r = 2 To srcLastRow
        gMigrationRow = r
        If Len(Trim$(srcWs.Cells(r, anchorCol).Value & "")) > 0 Then
            Dim destRow As ListRow
            Set destRow = NextTableRow(destTbl)
            For i = 1 To n
                If srcCols(i) > 0 And destCols(i) > 0 Then
                    destWs.Cells(destRow.Range.Row, destCols(i)).Value = srcWs.Cells(r, srcCols(i)).Value
                End If
            Next i
            modTasks.FinishArchiveRow destWs, destRow.Range.Row
            migrated = migrated + 1
        End If
    Next r

    gMigrationLog = gMigrationLog & "Archive: migrated " & migrated & " row(s)."
    If Len(missing) > 0 Then gMigrationLog = gMigrationLog & " (old file was missing: " & Left$(missing, Len(missing) - 2) & ")"
    gMigrationLog = gMigrationLog & vbCrLf
End Sub

Public Sub MigrateCampaignsAndArchive(srcWb As Workbook)
    Dim srcWs As Worksheet
    Set srcWs = TrySheet(srcWb, "Campaigns")
    If srcWs Is Nothing Then
        gMigrationLog = gMigrationLog & "Campaigns: SKIPPED (sheet not found)." & vbCrLf
        Exit Sub
    End If
    Dim destWs As Worksheet
    Set destWs = ThisWorkbook.Sheets("Campaigns")

    Dim nameCol As Long, startCol As Long, endCol As Long, statusCol As Long, notesCol As Long, keyCol As Long
    nameCol = FindColumnByHeaderText(srcWs, 2, "Campaign Name")
    startCol = FindColumnByHeaderText(srcWs, 2, "Start Date")
    endCol = FindColumnByHeaderText(srcWs, 2, "End Date")
    statusCol = FindColumnByHeaderText(srcWs, 2, "Status")
    notesCol = FindColumnByHeaderText(srcWs, 2, "Notes")
    keyCol = FindColumnByHeaderText(srcWs, 2, "CampaignKey")

    If nameCol = 0 Or keyCol = 0 Then
        gMigrationLog = gMigrationLog & "Campaigns: SKIPPED (couldn't find the Campaign Name/CampaignKey columns)." & vbCrLf
        Exit Sub
    End If

    Dim destNameCol As Long, destStartCol As Long, destEndCol As Long, destStatusCol As Long, destNotesCol As Long, destKeyCol As Long
    destNameCol = FindColumnByHeaderText(destWs, 2, "Campaign Name")
    destStartCol = FindColumnByHeaderText(destWs, 2, "Start Date")
    destEndCol = FindColumnByHeaderText(destWs, 2, "End Date")
    destStatusCol = FindColumnByHeaderText(destWs, 2, "Status")
    destNotesCol = FindColumnByHeaderText(destWs, 2, "Notes")
    destKeyCol = FindColumnByHeaderText(destWs, 2, "CampaignKey")

    ' Fixed 20-slot layout on both sides (modCampaigns.bas's own header
    ' comment) -- these mirror its Private FIRST_DATA_ROW/SLOT_SIZE/
    ' NUM_SLOTS exactly (not exposed publicly there, so duplicated here;
    ' confirmed unchanged since the original handoff commit).
    Const FIRST_MAIN_ROW As Long = 3
    Const SLOT_SIZE As Long = 10
    Const NUM_SLOTS As Long = 20

    Dim destSlot As Long, overflow As Long
    Dim s As Long, srcRow As Long
    For s = 0 To NUM_SLOTS - 1
        srcRow = FIRST_MAIN_ROW + s * SLOT_SIZE
        If Len(Trim$(srcWs.Cells(srcRow, nameCol).Value & "")) > 0 Then
            If destSlot >= NUM_SLOTS Then
                overflow = overflow + 1
            Else
                Dim destRow As Long
                destRow = FIRST_MAIN_ROW + destSlot * SLOT_SIZE
                destWs.Cells(destRow, destNameCol).Value = srcWs.Cells(srcRow, nameCol).Value
                If destStartCol > 0 And startCol > 0 Then destWs.Cells(destRow, destStartCol).Value = srcWs.Cells(srcRow, startCol).Value
                If destEndCol > 0 And endCol > 0 Then destWs.Cells(destRow, destEndCol).Value = srcWs.Cells(srcRow, endCol).Value
                If destStatusCol > 0 And statusCol > 0 Then destWs.Cells(destRow, destStatusCol).Value = srcWs.Cells(srcRow, statusCol).Value
                If destNotesCol > 0 And notesCol > 0 Then destWs.Cells(destRow, destNotesCol).Value = srcWs.Cells(srcRow, notesCol).Value
                destWs.Cells(destRow, destKeyCol).Value = srcWs.Cells(srcRow, keyCol).Value
                destSlot = destSlot + 1
            End If
        End If
    Next s

    ' Reuses the already-tested engine (also applies real button styling
    ' to Up/Down, per this session's own fix, and rebuilds the dropdown
    ' mirror) instead of duplicating any of that here.
    modCampaigns.CompactCampaignSlots

    gMigrationLog = gMigrationLog & "Campaigns: migrated " & destSlot & " active/planned campaign(s)."
    If overflow > 0 Then gMigrationLog = gMigrationLog & " WARNING: " & overflow & " more campaign(s) in the old file didn't fit in the 20 slots and were NOT migrated."
    gMigrationLog = gMigrationLog & vbCrLf

End Sub

Public Sub MigrateCampaignsArchive(srcWb As Workbook)
    Dim srcWs As Worksheet
    Set srcWs = TrySheet(srcWb, "Campaigns Archive")
    If srcWs Is Nothing Then
        gMigrationLog = gMigrationLog & "Campaigns Archive: SKIPPED (sheet not found)." & vbCrLf
        Exit Sub
    End If
    Dim destWs As Worksheet
    Set destWs = ThisWorkbook.Sheets("Campaigns Archive")

    Dim labels() As Variant
    labels = Array("Campaign Name", "Start Date", "End Date", "Status", "Progress", "Notes", "Completed At", "CampaignKey")

    Dim anchorCol As Long
    anchorCol = FindColumnByHeaderText(srcWs, 1, "Campaign Name")
    If anchorCol = 0 Then
        gMigrationLog = gMigrationLog & "Campaigns Archive: SKIPPED (couldn't find 'Campaign Name' column)." & vbCrLf
        Exit Sub
    End If

    Dim srcLastRow As Long
    srcLastRow = modHierarchy.LastHierarchyRow(srcWs, anchorCol, 2)

    Dim destTbl As ListObject
    Set destTbl = destWs.ListObjects("CampaignsArchiveTable")

    Dim n As Long
    n = UBound(labels) - LBound(labels) + 1
    Dim srcCols() As Long, destCols() As Long
    ReDim srcCols(1 To n)
    ReDim destCols(1 To n)
    Dim i As Long, missing As String
    For i = 1 To n
        srcCols(i) = FindColumnByHeaderText(srcWs, 1, CStr(labels(LBound(labels) + i - 1)))
        destCols(i) = FindDestColumn(destWs, CStr(labels(LBound(labels) + i - 1)))
        If srcCols(i) = 0 And Not IsNewInV4(CStr(labels(LBound(labels) + i - 1))) Then missing = missing & CStr(labels(LBound(labels) + i - 1)) & ", "
    Next i

    Dim r As Long, migrated As Long
    For r = 2 To srcLastRow
        gMigrationRow = r
        If Len(Trim$(srcWs.Cells(r, anchorCol).Value & "")) > 0 Then
            Dim destRow As ListRow
            Set destRow = NextTableRow(destTbl)
            For i = 1 To n
                If srcCols(i) > 0 And destCols(i) > 0 Then
                    destWs.Cells(destRow.Range.Row, destCols(i)).Value = srcWs.Cells(r, srcCols(i)).Value
                End If
            Next i
            migrated = migrated + 1
        End If
    Next r

    gMigrationLog = gMigrationLog & "Campaigns Archive: migrated " & migrated & " row(s)."
    If Len(missing) > 0 Then gMigrationLog = gMigrationLog & " (old file was missing: " & Left$(missing, Len(missing) - 2) & ")"
    gMigrationLog = gMigrationLog & vbCrLf
End Sub

Public Sub MigrateRecurringTasks(srcWb As Workbook)
    Dim srcWs As Worksheet
    Set srcWs = TrySheet(srcWb, "Recurring Tasks")
    If srcWs Is Nothing Then
        gMigrationLog = gMigrationLog & "Recurring Tasks: SKIPPED (sheet not found)." & vbCrLf
        Exit Sub
    End If
    Dim destWs As Worksheet
    Set destWs = ThisWorkbook.Sheets("Recurring Tasks")

    Dim keyCol As Long
    keyCol = FindColumnByHeaderText(srcWs, 1, "RecurringKey")
    If keyCol = 0 Then
        gMigrationLog = gMigrationLog & "Recurring Tasks: SKIPPED (couldn't find the RecurringKey column)." & vbCrLf
        Exit Sub
    End If

    Dim srcLastRow As Long
    srcLastRow = modHierarchy.LastHierarchyRow(srcWs, keyCol, 2)

    Dim srcRowCount As Long, r As Long
    For r = 2 To srcLastRow
        gMigrationRow = r
        If Len(Trim$(srcWs.Cells(r, keyCol).Value & "")) > 0 Then srcRowCount = srcRowCount + 1
    Next r

    If srcRowCount = 0 Then
        gMigrationLog = gMigrationLog & "Recurring Tasks: nothing to migrate." & vbCrLf
        Exit Sub
    End If

    ' Recurring Tasks has no row-extender of its own -- its 100-row
    ' allowance is meant to comfortably fit a curated template list
    ' (sheets/recurring_tasks.py module docstring). Migrate what fits and
    ' warn loudly rather than silently overflowing into whatever sits
    ' below row 101.
    Dim destCap As Long
    destCap = 100
    Dim toMigrate As Long, overflow As Long
    toMigrate = srcRowCount
    If toMigrate > destCap Then
        overflow = toMigrate - destCap
        toMigrate = destCap
    End If

    Dim labels() As Variant
    labels = Array("Status", "Task", "Project / Category", "Linked Campaign", "Priority", "Recurrence", _
                    "Task Start On", "Time", "If Missed", "Deadline Offset (days)", "Notes", "RecurringKey", "RecurringParentKey", _
                    "SiblingOrder", "LinkedCampaignKey", "Last Generated")

    Dim n As Long
    n = UBound(labels) - LBound(labels) + 1
    Dim srcCols() As Long, destCols() As Long
    ReDim srcCols(1 To n)
    ReDim destCols(1 To n)
    Dim i As Long, missing As String
    For i = 1 To n
        srcCols(i) = FindColumnByHeaderText(srcWs, 1, CStr(labels(LBound(labels) + i - 1)))
        destCols(i) = FindDestColumn(destWs, CStr(labels(LBound(labels) + i - 1)))
        If srcCols(i) = 0 And Not IsNewInV4(CStr(labels(LBound(labels) + i - 1))) Then missing = missing & CStr(labels(LBound(labels) + i - 1)) & ", "
    Next i

    Dim destRow As Long, copied As Long
    destRow = 2
    For r = 2 To srcLastRow
        gMigrationRow = r
        If Len(Trim$(srcWs.Cells(r, keyCol).Value & "")) > 0 Then
            If copied < toMigrate Then
                For i = 1 To n
                    If srcCols(i) > 0 And destCols(i) > 0 Then
                        destWs.Cells(destRow, destCols(i)).Value = srcWs.Cells(r, srcCols(i)).Value
                    End If
                Next i
                destRow = destRow + 1
                copied = copied + 1
            End If
        End If
    Next r

    ' Recurring Tasks' real column map (Public since V4 -- this module
    ' used to keep its own hard-coded copy, which went stale the moment
    ' Time/If Missed were inserted, PLAN.md §22.9).
    modHierarchy.RenumberHierarchy destWs, modRecurringTasks.RTColMap()
    modRecurringTasks.RefreshButtonStyling destWs
    modRecurringTasks.RefreshLinkedCampaignLabels destWs

    ' V4 (PLAN.md §22.3): a subtask's own Recurrence used to be ignored;
    ' now it decides which days that subtask is generated on. Say so, so
    ' an old leftover value doesn't silently start filtering.
    Dim ownSchedule As Long, rr As Long
    Dim rtMap As modHierarchy.HierarchyColumnMap
    rtMap = modRecurringTasks.RTColMap()
    Dim recCol As Long
    recCol = FindColumnByHeaderText(destWs, 1, "Recurrence")
    For rr = 2 To destRow - 1
        If Len(Trim$(destWs.Cells(rr, rtMap.ColParentKey).Value & "")) > 0 And recCol > 0 Then
            If Len(Trim$(destWs.Cells(rr, recCol).Value & "")) > 0 Then ownSchedule = ownSchedule + 1
        End If
    Next rr

    gMigrationLog = gMigrationLog & "Recurring Tasks: migrated " & copied & " row(s)."
    If ownSchedule > 0 Then gMigrationLog = gMigrationLog & " NOTE: " & ownSchedule & _
        " subtask(s) have their own Recurrence -- since V4 that decides which days they're generated on " & _
        "(blank = same as parent). Check them on the Recurring Tasks sheet."
    If overflow > 0 Then gMigrationLog = gMigrationLog & " WARNING: " & overflow & " more row(s) didn't fit in the 100-row allowance and were NOT migrated."
    If Len(missing) > 0 Then gMigrationLog = gMigrationLog & " (old file was missing: " & Left$(missing, Len(missing) - 2) & ")"
    gMigrationLog = gMigrationLog & vbCrLf
End Sub

' Shared engine for the plain flat-table sheets: Login Information, Quick
' Links, and Schedule Data's two side-by-side tables (called once per
' table -- Schedule Data's own header row has both tables' labels on it,
' so FindColumnByHeaderText finds either regardless of which table it
' belongs to). `labels`'s FIRST entry is used as the row-presence anchor,
' same "first column of the row" convention used project-wide.
Public Sub MigrateFlatTable(srcWb As Workbook, sheetName As String, tableName As String, labels As Variant)
    Dim srcWs As Worksheet
    Set srcWs = TrySheet(srcWb, sheetName)
    If srcWs Is Nothing Then
        gMigrationLog = gMigrationLog & sheetName & " (" & tableName & "): SKIPPED (sheet not found)." & vbCrLf
        Exit Sub
    End If
    Dim destWs As Worksheet
    Set destWs = ThisWorkbook.Sheets(sheetName)

    Dim anchorLabel As String
    anchorLabel = CStr(labels(LBound(labels)))
    Dim anchorCol As Long
    anchorCol = FindColumnByHeaderText(srcWs, 1, anchorLabel)
    If anchorCol = 0 Then
        gMigrationLog = gMigrationLog & sheetName & " (" & tableName & "): SKIPPED (couldn't find the '" & anchorLabel & "' column)." & vbCrLf
        Exit Sub
    End If

    Dim srcLastRow As Long
    srcLastRow = modHierarchy.LastHierarchyRow(srcWs, anchorCol, 2)

    Dim destTbl As ListObject
    Set destTbl = destWs.ListObjects(tableName)

    Dim n As Long
    n = UBound(labels) - LBound(labels) + 1
    Dim srcCols() As Long, destCols() As Long
    ReDim srcCols(1 To n)
    ReDim destCols(1 To n)
    Dim i As Long, missing As String
    For i = 1 To n
        srcCols(i) = FindColumnByHeaderText(srcWs, 1, CStr(labels(LBound(labels) + i - 1)))
        destCols(i) = FindDestColumn(destWs, CStr(labels(LBound(labels) + i - 1)))
        If srcCols(i) = 0 And Not IsNewInV4(CStr(labels(LBound(labels) + i - 1))) Then missing = missing & CStr(labels(LBound(labels) + i - 1)) & ", "
    Next i

    Dim r As Long, migrated As Long
    For r = 2 To srcLastRow
        gMigrationRow = r
        If Len(Trim$(srcWs.Cells(r, anchorCol).Value & "")) > 0 Then
            Dim destRow As ListRow
            Set destRow = NextTableRow(destTbl)
            For i = 1 To n
                If srcCols(i) > 0 And destCols(i) > 0 Then
                    destWs.Cells(destRow.Range.Row, destCols(i)).Value = srcWs.Cells(r, srcCols(i)).Value
                End If
            Next i
            migrated = migrated + 1
        End If
    Next r

    gMigrationLog = gMigrationLog & sheetName & " (" & tableName & "): migrated " & migrated & " row(s)."
    If Len(missing) > 0 Then gMigrationLog = gMigrationLog & " (old file was missing: " & Left$(missing, Len(missing) - 2) & ")"
    gMigrationLog = gMigrationLog & vbCrLf
End Sub
