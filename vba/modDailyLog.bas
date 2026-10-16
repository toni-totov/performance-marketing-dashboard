Attribute VB_Name = "modDailyLog"
Option Explicit

' Daily Log sheet logic (PLAN.md §22.6, V4, employee only). A read-only
' week view over Archive: one section per day, each listing what was
' completed (or skipped) that day in completion-time order. Archive itself
' stays the single flat log the Dashboard counts from -- this only
' re-groups it, so nothing here ever writes to Archive.
'
' Every public entry point is error-handled (CLAUDE.md: never let an
' unhandled error surface a blocking modal) and every display glyph is
' built with ChrW (VBA source is ANSI -- a literal ◀/✓/› would corrupt to
' "?" on compile, CLAUDE.md). Day names are a fixed English list rather
' than Format(d, "dddd"), which follows the machine's locale.
'
' CopyForStandup is the MsgBox-showing button wrapper; CopyForStandupCore
' (copy + return the text, no modal) is what tests call -- same split as
' frmNotes' ShowForCell/PrepareForCell.

Private Const SHEET_NAME As String = "Daily Log"
Private Const BODY_FIRST_ROW As Long = 4
Private Const BODY_LAST_COL As Long = 4
Private Const MAX_BODY_ROW As Long = 3000

' Tasks' own layout (employee) -- only what "Today" in the stand-up text
' needs. Mirrors modTasks.bas's private constants.
Private Const TASKS_COL_TASK As Long = 5
Private Const TASKS_COL_DEADLINE As Long = 11
Private Const TASKS_COL_TASK_KEY As Long = 17
Private Const TASKS_COL_PARENT_KEY As Long = 18
Private Const TASKS_FIRST_DATA_ROW As Long = 2

Private Function LogSheet() As Worksheet
    Set LogSheet = ThisWorkbook.Sheets(SHEET_NAME)
End Function

' ============================================================================
' Week navigation
' ============================================================================

Public Function WeekStartOf(d As Date) As Date
    WeekStartOf = Int(d) - (Weekday(d, vbMonday) - 1)
End Function

Private Function CurrentWeekStart() As Date
    Dim v As Variant
    v = LogSheet().Range("DailyLogWeekStart").Value
    If IsDate(v) Then
        CurrentWeekStart = WeekStartOf(CDate(v))
    ElseIf IsNumeric(v) And Not IsEmpty(v) Then
        CurrentWeekStart = WeekStartOf(CDate(CDbl(v)))
    Else
        CurrentWeekStart = WeekStartOf(Date)
    End If
End Function

Public Function GoToWeek(weekStart As Date) As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = LogSheet()
    ws.Unprotect
    ws.Range("DailyLogWeekStart").Value = WeekStartOf(weekStart)
    GoToWeek = RenderWeek()
    Exit Function
ErrHandler:
    GoToWeek = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Sub GoThisWeek()
    GoToWeek WeekStartOf(Date)
End Sub

Public Sub GoPrevWeek()
    GoToWeek CurrentWeekStart() - 7
End Sub

Public Sub GoNextWeek()
    GoToWeek CurrentWeekStart() + 7
End Sub

' ============================================================================
' Reading Archive / Tasks
' ============================================================================

Private Function DayName(d As Date) As String
    DayName = Choose(Weekday(d, vbMonday), "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
End Function

Private Function DateText(d As Date, Optional withYear As Boolean = True) As String
    DateText = Format(Day(d), "00") & "." & Format(Month(d), "00")
    If withYear Then DateText = DateText & "." & Year(d)
End Function

Private Function TimeText(t As Double) As String
    Dim mins As Long
    mins = CLng(Round((t - Int(t)) * 1440, 0))
    If mins >= 1440 Then mins = 1439
    TimeText = Format(mins \ 60, "00") & ":" & Format(mins Mod 60, "00")
End Function

' A key cell's value as comparable text ("12" for 12, 12.0 or "12").
Private Function KeyText(v As Variant) As String
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(v & "")) = 0 Then Exit Function
    If IsNumeric(v) Then
        KeyText = CStr(CDbl(v))
    Else
        KeyText = Trim$(v & "")
    End If
End Function

' TaskKey -> task name, across Archive AND Tasks: a subtask completed on
' its own has its parent still open in Tasks; one completed together with
' its parent finds the parent in Archive.
Private Function TaskNamesByKey() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim aws As Worksheet
    Set aws = ThisWorkbook.Sheets("Archive")
    Dim cKey As Long, cTask As Long
    cKey = modTasks.ArchiveColumn("TaskKey")
    cTask = modTasks.ArchiveColumn("Task")
    Dim lastA As Long
    lastA = ArchiveLastRow()
    Dim r As Long
    If cKey > 0 And cTask > 0 Then
        For r = 2 To lastA
            Dim k As String
            k = KeyText(aws.Cells(r, cKey).Value)
            If Len(k) > 0 Then d(k) = aws.Cells(r, cTask).Value & ""
        Next r
    End If

    Dim tws As Worksheet
    Set tws = ThisWorkbook.Sheets("Tasks")
    Dim lastT As Long
    lastT = modHierarchy.LastHierarchyRow(tws, TASKS_COL_TASK_KEY, TASKS_FIRST_DATA_ROW)
    For r = TASKS_FIRST_DATA_ROW To lastT
        k = KeyText(tws.Cells(r, TASKS_COL_TASK_KEY).Value)
        If Len(k) > 0 Then d(k) = tws.Cells(r, TASKS_COL_TASK).Value & ""
    Next r

    Set TaskNamesByKey = d
End Function

Private Function TaskPath(taskName As String, parentKey As String, names As Object) As String
    If Len(parentKey) > 0 Then
        If names.Exists(parentKey) Then
            TaskPath = names(parentKey) & " " & ChrW(8250) & " " & taskName
            Exit Function
        End If
    End If
    TaskPath = taskName
End Function

' Archive is a Table that starts with one blank template row -- the last
' row with a real Task, scanned manually (hidden columns, CLAUDE.md).
Private Function ArchiveLastRow() As Long
    Dim aws As Worksheet
    Set aws = ThisWorkbook.Sheets("Archive")
    Dim tbl As ListObject
    Set tbl = aws.ListObjects("ArchiveTable")
    ArchiveLastRow = tbl.Range.Row + tbl.Range.Rows.Count - 1
End Function

' Every Archive row completed in [fromDate, toDate), sorted by completion
' time. Parallel arrays, 1-based, sized n.
Private Sub CollectArchived(fromDate As Date, toDate As Date, names As Object, _
                            ByRef whenArr() As Double, ByRef textArr() As String, ByRef projArr() As String, _
                            ByRef skippedArr() As Boolean, ByRef n As Long)
    n = 0
    Dim aws As Worksheet
    Set aws = ThisWorkbook.Sheets("Archive")
    Dim cAt As Long, cTask As Long, cProj As Long, cStatus As Long, cParent As Long
    cAt = modTasks.ArchiveColumn("Completed At")
    cTask = modTasks.ArchiveColumn("Task")
    cProj = modTasks.ArchiveColumn("Project / Category")
    cStatus = modTasks.ArchiveColumn("Status")
    cParent = modTasks.ArchiveColumn("ParentKey")
    If cAt = 0 Or cTask = 0 Then Exit Sub

    Dim lastA As Long
    lastA = ArchiveLastRow()
    ReDim whenArr(1 To lastA)
    ReDim textArr(1 To lastA)
    ReDim projArr(1 To lastA)
    ReDim skippedArr(1 To lastA)

    Dim r As Long
    For r = 2 To lastA
        Dim v As Variant
        v = aws.Cells(r, cAt).Value
        If IsDate(v) And Len(Trim$(aws.Cells(r, cTask).Value & "")) > 0 Then
            Dim t As Double
            t = CDbl(CDate(v))
            If t >= CDbl(fromDate) And t < CDbl(toDate) Then
                n = n + 1
                whenArr(n) = t
                Dim pk As String
                If cParent > 0 Then pk = KeyText(aws.Cells(r, cParent).Value) Else pk = ""
                textArr(n) = TaskPath(aws.Cells(r, cTask).Value & "", pk, names)
                If cProj > 0 Then projArr(n) = aws.Cells(r, cProj).Value & ""
                If cStatus > 0 Then skippedArr(n) = (Trim$(aws.Cells(r, cStatus).Value & "") = "Skipped")
            End If
        End If
    Next r

    ' insertion sort by time (stable -- same-moment rows of one cascade
    ' keep Archive's own depth-first order, parent before child)
    Dim i As Long, j As Long
    For i = 2 To n
        Dim tw As Double, tt As String, tp As String, ts As Boolean
        tw = whenArr(i): tt = textArr(i): tp = projArr(i): ts = skippedArr(i)
        j = i - 1
        Do While j >= 1
            If whenArr(j) <= tw Then Exit Do
            whenArr(j + 1) = whenArr(j): textArr(j + 1) = textArr(j)
            projArr(j + 1) = projArr(j): skippedArr(j + 1) = skippedArr(j)
            j = j - 1
        Loop
        whenArr(j + 1) = tw: textArr(j + 1) = tt: projArr(j + 1) = tp: skippedArr(j + 1) = ts
    Next i
End Sub

' ============================================================================
' Rendering
' ============================================================================

Private Function WeekLabel(weekStart As Date) As String
    Dim wk As Long
    wk = DatePart("ww", weekStart, vbMonday, vbFirstFourDays)
    WeekLabel = "Week " & wk & "  " & ChrW(183) & "  " & DateText(weekStart, False) & " " & ChrW(8211) & " " & DateText(weekStart + 6)
End Function

Public Function RenderWeek() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = LogSheet()
    Dim prevUpdating As Boolean
    prevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim weekStart As Date
    weekStart = CurrentWeekStart()

    ws.Unprotect
    ws.Range("DailyLogWeekStart").Value = weekStart
    ws.Range("B1").Value = WeekLabel(weekStart)

    With ws.Range(ws.Cells(BODY_FIRST_ROW, 1), ws.Cells(MAX_BODY_ROW, BODY_LAST_COL))
        .UnMerge
        .Clear
        .Font.Name = modDesignSystem.FONT_NAME
        .Font.Size = modDesignSystem.FONT_SIZE_BODY
    End With
    ws.Rows(BODY_FIRST_ROW & ":" & MAX_BODY_ROW).RowHeight = 19

    Dim names As Object
    Set names = TaskNamesByKey()
    Dim whenArr() As Double, textArr() As String, projArr() As String, skippedArr() As Boolean, n As Long
    CollectArchived weekStart, weekStart + 7, names, whenArr, textArr, projArr, skippedArr, n

    Dim r As Long
    r = BODY_FIRST_ROW
    Dim d As Long, i As Long
    For d = 0 To 6
        Dim dayDate As Date
        dayDate = weekStart + d
        Dim doneCount As Long, skipCount As Long
        doneCount = 0: skipCount = 0
        For i = 1 To n
            If Int(whenArr(i)) = CDbl(dayDate) Then
                If skippedArr(i) Then skipCount = skipCount + 1 Else doneCount = doneCount + 1
            End If
        Next i

        If d < 5 Or doneCount + skipCount > 0 Then
            ' Day header
            Dim title As String
            title = DayName(dayDate) & " " & ChrW(8211) & " " & DateText(dayDate)
            If dayDate = Date Then title = title & "  (today)"
            title = title & "      " & doneCount & " done"
            If skipCount > 0 Then title = title & "  " & ChrW(183) & "  " & skipCount & " skipped"
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, BODY_LAST_COL))
                .Merge
                .Value = title
                .Interior.Color = IIf(dayDate = Date, modDesignSystem.COLOR_CATEGORICAL_1, modDesignSystem.COLOR_BRAND_ACCENT)
                .Font.Color = modDesignSystem.COLOR_BRAND_ACCENT_TEXT
                .Font.Bold = True
                .Font.Size = 12
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .IndentLevel = 1
            End With
            ws.Rows(r).RowHeight = 26
            r = r + 1

            If doneCount + skipCount = 0 Then
                With ws.Range(ws.Cells(r, 1), ws.Cells(r, BODY_LAST_COL))
                    .Merge
                    .Value = "Nothing logged"
                    .Font.Italic = True
                    .Font.Color = modDesignSystem.COLOR_INK_MUTED
                    .IndentLevel = 1
                End With
                r = r + 1
            Else
                For i = 1 To n
                    If Int(whenArr(i)) = CDbl(dayDate) Then
                        ws.Cells(r, 1).Value = "'" & TimeText(whenArr(i))
                        ws.Cells(r, 1).HorizontalAlignment = xlCenter
                        ws.Cells(r, 2).Value = textArr(i)
                        ws.Cells(r, 3).Value = projArr(i)
                        If skippedArr(i) Then
                            ws.Cells(r, 4).Value = ChrW(10556) & " Skipped"
                            With ws.Range(ws.Cells(r, 1), ws.Cells(r, BODY_LAST_COL)).Font
                                .Italic = True
                                .Color = modDesignSystem.COLOR_INK_MUTED
                            End With
                        Else
                            ws.Cells(r, 4).Value = ChrW(10003) & " Done"
                            ws.Cells(r, 4).Font.Color = modDesignSystem.COLOR_STATUS_GOOD
                            ws.Cells(r, 4).Font.Bold = True
                        End If
                        With ws.Range(ws.Cells(r, 1), ws.Cells(r, BODY_LAST_COL)).Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .Weight = xlHairline
                            .Color = modDesignSystem.COLOR_GRIDLINE
                        End With
                        r = r + 1
                    End If
                Next i
            End If
            r = r + 1   ' spacer between days
        End If
    Next d

    ws.Protect UserInterfaceOnly:=True
    Application.ScreenUpdating = prevUpdating
    RenderWeek = "OK " & n
    Exit Function
ErrHandler:
    RenderWeek = "ERROR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    Application.ScreenUpdating = prevUpdating
End Function

' ============================================================================
' Stand-up text
' ============================================================================

Public Function PreviousWorkday(d As Date) As Date
    Dim p As Date
    p = Int(d) - 1
    Do While Weekday(p, vbMonday) > 5
        p = p - 1
    Loop
    PreviousWorkday = p
End Function

' Plain text for a daily stand-up: what got done on the previous working
' day (Monday looks back to Friday), plus what's due today. Pure -- no
' clipboard, no MsgBox.
Public Function BuildStandupText(forDate As Date) As String
    On Error GoTo ErrHandler
    Dim names As Object
    Set names = TaskNamesByKey()
    Dim prevDay As Date
    prevDay = PreviousWorkday(forDate)

    Dim whenArr() As Double, textArr() As String, projArr() As String, skippedArr() As Boolean, n As Long
    CollectArchived prevDay, prevDay + 1, names, whenArr, textArr, projArr, skippedArr, n

    Dim s As String, i As Long
    Dim label As String
    If prevDay = Int(forDate) - 1 Then label = "Yesterday" Else label = DayName(prevDay)
    s = label & " (" & DateText(prevDay, False) & "):" & vbCrLf
    Dim doneAny As Boolean, skipped As String
    For i = 1 To n
        If skippedArr(i) Then
            If Len(skipped) > 0 Then skipped = skipped & ", "
            skipped = skipped & textArr(i)
        Else
            s = s & "- " & textArr(i) & vbCrLf
            doneAny = True
        End If
    Next i
    If Not doneAny Then s = s & "- (nothing logged)" & vbCrLf
    If Len(skipped) > 0 Then s = s & "Skipped: " & skipped & vbCrLf

    ' Today: every open Tasks row due today, by deadline time.
    s = s & vbCrLf & "Today (" & DateText(forDate, False) & "):" & vbCrLf
    Dim tws As Worksheet
    Set tws = ThisWorkbook.Sheets("Tasks")
    Dim lastT As Long
    lastT = modHierarchy.LastHierarchyRow(tws, TASKS_COL_TASK_KEY, TASKS_FIRST_DATA_ROW)
    Dim tm() As Double, tx() As String, m As Long, r As Long
    If lastT >= TASKS_FIRST_DATA_ROW Then
        ReDim tm(1 To lastT)
        ReDim tx(1 To lastT)
    End If
    For r = TASKS_FIRST_DATA_ROW To lastT
        If Len(KeyText(tws.Cells(r, TASKS_COL_TASK_KEY).Value)) > 0 Then
            Dim dl As Variant
            dl = tws.Cells(r, TASKS_COL_DEADLINE).Value
            If IsDate(dl) Then
                If Int(CDbl(CDate(dl))) = CDbl(Int(forDate)) Then
                    m = m + 1
                    tm(m) = CDbl(CDate(dl))
                    tx(m) = TaskPath(tws.Cells(r, TASKS_COL_TASK).Value & "", KeyText(tws.Cells(r, TASKS_COL_PARENT_KEY).Value), names)
                End If
            End If
        End If
    Next r
    Dim j As Long
    For i = 2 To m
        Dim kt As Double, ks As String
        kt = tm(i): ks = tx(i)
        j = i - 1
        Do While j >= 1
            If tm(j) <= kt Then Exit Do
            tm(j + 1) = tm(j): tx(j + 1) = tx(j)
            j = j - 1
        Loop
        tm(j + 1) = kt: tx(j + 1) = ks
    Next i
    For i = 1 To m
        s = s & "- " & TimeText(tm(i)) & "  " & tx(i) & vbCrLf
    Next i
    If m = 0 Then s = s & "- (nothing due)" & vbCrLf

    BuildStandupText = s
    Exit Function
ErrHandler:
    BuildStandupText = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Puts text on the Windows clipboard. The htmlfile route needs no
' library reference; the MSForms DataObject (late-bound by CLSID) is the
' fallback. Returns False if neither worked.
Public Function SetClipboardText(txt As String) As Boolean
    On Error GoTo TryForms
    Dim html As Object
    Set html = CreateObject("htmlfile")
    html.ParentWindow.ClipboardData.SetData "text", txt
    SetClipboardText = True
    Exit Function
TryForms:
    On Error GoTo Failed
    Dim dobj As Object
    Set dobj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dobj.SetText txt
    dobj.PutInClipboard
    SetClipboardText = True
    Exit Function
Failed:
    SetClipboardText = False
End Function

' Returns "COPIED" & vbLf & text, or "NOT_COPIED" & vbLf & text.
Public Function CopyForStandupCore(forDate As Date) As String
    On Error GoTo ErrHandler
    Dim txt As String
    txt = BuildStandupText(forDate)
    If SetClipboardText(txt) Then
        CopyForStandupCore = "COPIED" & vbLf & txt
    Else
        CopyForStandupCore = "NOT_COPIED" & vbLf & txt
    End If
    Exit Function
ErrHandler:
    CopyForStandupCore = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Sub CopyForStandup()
    On Error GoTo ErrHandler
    Dim res As String
    res = CopyForStandupCore(Date)
    Dim sep As Long
    sep = InStr(res, vbLf)
    If Left$(res, 6) = "COPIED" Then
        MsgBox "Copied to the clipboard -- paste it into your meeting notes or chat:" & vbCrLf & vbCrLf & Mid$(res, sep + 1), vbInformation, "Stand-up"
    ElseIf sep > 0 Then
        MsgBox "Couldn't reach the clipboard on this machine -- here's the text:" & vbCrLf & vbCrLf & Mid$(res, sep + 1), vbExclamation, "Stand-up"
    Else
        MsgBox res, vbExclamation, "Stand-up"
    End If
    Exit Sub
ErrHandler:
End Sub
