Attribute VB_Name = "modToday"
Option Explicit

' Dashboard "Today" panel (PLAN.md §22.7, V4, employee only): today's open
' Tasks (deadline today) and today's Schedule meetings, merged into one
' list in time order, plus an overdue count. The frame (title, header,
' merged N:P item cells) is built by sheets/dashboard.py; this only fills
' values and per-row styling. Redrawn on Workbook_Open, on every Dashboard
' visit and by Refresh Now.

Private Const COL_TIME As Long = 13      ' M
Private Const COL_ITEM As Long = 14      ' N (merged N:P)
Private Const COL_TYPE As Long = 17      ' Q
Private Const FIRST_ITEM_ROW As Long = 8
Private Const LAST_ITEM_ROW As Long = 19
Private Const FOOTER_ROW As Long = 20

' Tasks (employee layout) -- mirrors modTasks.bas's private constants.
Private Const TASKS_COL_TASK As Long = 5
Private Const TASKS_COL_DEADLINE As Long = 11
Private Const TASKS_COL_STATUS As Long = 12
Private Const TASKS_COL_TASK_KEY As Long = 17
Private Const TASKS_COL_PARENT_KEY As Long = 18
Private Const TASKS_FIRST_DATA_ROW As Long = 2

' Schedule Data's MeetingsTable -- mirrors modSchedule.bas.
Private Const MT_COL_NAME As Long = 4
Private Const MT_COL_START As Long = 5
Private Const MT_COL_END As Long = 6

Private Function TimeText(t As Double) As String
    Dim mins As Long
    mins = CLng(Round((t - Int(t)) * 1440, 0))
    If mins >= 1440 Then mins = 1439
    TimeText = Format(mins \ 60, "00") & ":" & Format(mins Mod 60, "00")
End Function

Private Function KeyText(v As Variant) As String
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(v & "")) = 0 Then Exit Function
    If IsNumeric(v) Then KeyText = CStr(CDbl(v)) Else KeyText = Trim$(v & "")
End Function

Public Function RenderTodayPanel() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    Dim today As Date
    today = Date
    Dim nowFrac As Double
    nowFrac = CDbl(Now) - Int(CDbl(Now))

    ' --- collect ---
    Dim tm() As Double, tmEnd() As Double, itemText() As String, kind() As String
    Dim n As Long
    ReDim tm(1 To 2000): ReDim tmEnd(1 To 2000): ReDim itemText(1 To 2000): ReDim kind(1 To 2000)

    Dim tws As Worksheet
    Set tws = ThisWorkbook.Sheets("Tasks")
    Dim lastT As Long, r As Long
    lastT = modHierarchy.LastHierarchyRow(tws, TASKS_COL_TASK_KEY, TASKS_FIRST_DATA_ROW)

    Dim names As Object
    Set names = CreateObject("Scripting.Dictionary")
    For r = TASKS_FIRST_DATA_ROW To lastT
        If Len(KeyText(tws.Cells(r, TASKS_COL_TASK_KEY).Value)) > 0 Then
            names(KeyText(tws.Cells(r, TASKS_COL_TASK_KEY).Value)) = tws.Cells(r, TASKS_COL_TASK).Value & ""
        End If
    Next r

    Dim overdue As Long
    For r = TASKS_FIRST_DATA_ROW To lastT
        If Len(KeyText(tws.Cells(r, TASKS_COL_TASK_KEY).Value)) > 0 Then
            Dim dl As Variant
            dl = tws.Cells(r, TASKS_COL_DEADLINE).Value
            If IsDate(dl) Then
                Dim dlv As Double
                dlv = CDbl(CDate(dl))
                If Int(dlv) < CDbl(today) Then
                    overdue = overdue + 1
                ElseIf Int(dlv) = CDbl(today) And n < 2000 Then
                    n = n + 1
                    tm(n) = dlv - Int(dlv)
                    tmEnd(n) = -1
                    Dim pk As String
                    pk = KeyText(tws.Cells(r, TASKS_COL_PARENT_KEY).Value)
                    itemText(n) = tws.Cells(r, TASKS_COL_TASK).Value & ""
                    If Len(pk) > 0 Then
                        If names.Exists(pk) Then itemText(n) = names(pk) & " " & ChrW(8250) & " " & itemText(n)
                    End If
                    kind(n) = "Task"
                    If Trim$(tws.Cells(r, TASKS_COL_STATUS).Value & "") = "Rescheduled" Then kind(n) = "Task (moved)"
                End If
            End If
        End If
    Next r

    Dim sws As Worksheet
    Set sws = ThisWorkbook.Sheets("Schedule Data")
    Dim mt As ListObject
    Set mt = sws.ListObjects("MeetingsTable")
    If Not mt.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To mt.ListRows.Count
            r = mt.DataBodyRange.Rows(i).Row
            If Len(Trim$(sws.Cells(r, MT_COL_NAME).Value & "")) > 0 Then
                If modSchedule.MeetingOccursOn(sws, r, today) And n < 2000 Then
                    n = n + 1
                    tm(n) = CDbl(sws.Cells(r, MT_COL_START).Value)
                    tm(n) = tm(n) - Int(tm(n))
                    If IsNumeric(sws.Cells(r, MT_COL_END).Value) And Len(sws.Cells(r, MT_COL_END).Value & "") > 0 Then
                        tmEnd(n) = CDbl(sws.Cells(r, MT_COL_END).Value)
                        tmEnd(n) = tmEnd(n) - Int(tmEnd(n))
                    Else
                        tmEnd(n) = -1
                    End If
                    itemText(n) = sws.Cells(r, MT_COL_NAME).Value & ""
                    kind(n) = "Meeting"
                End If
            End If
        Next i
    End If

    ' --- sort by time (stable insertion sort) ---
    Dim j As Long
    For i = 2 To n
        Dim a As Double, b As Double, c As String, k As String
        a = tm(i): b = tmEnd(i): c = itemText(i): k = kind(i)
        j = i - 1
        Do While j >= 1
            If tm(j) <= a Then Exit Do
            tm(j + 1) = tm(j): tmEnd(j + 1) = tmEnd(j): itemText(j + 1) = itemText(j): kind(j + 1) = kind(j)
            j = j - 1
        Loop
        tm(j + 1) = a: tmEnd(j + 1) = b: itemText(j + 1) = c: kind(j + 1) = k
    Next i

    ' --- render ---
    ws.Range("M6").Value = "Today " & ChrW(8211) & " " & _
        Choose(Weekday(today, vbMonday), "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday") & _
        " " & Format(Day(today), "00") & "." & Format(Month(today), "00")

    Dim body As Range
    Set body = ws.Range(ws.Cells(FIRST_ITEM_ROW, COL_TIME), ws.Cells(FOOTER_ROW, COL_TYPE))
    body.ClearContents
    body.Font.Italic = False
    body.Font.Bold = False
    body.Font.Color = modDesignSystem.COLOR_INK_PRIMARY

    Dim capacity As Long
    capacity = LAST_ITEM_ROW - FIRST_ITEM_ROW + 1
    Dim shown As Long
    shown = n
    If shown > capacity Then shown = capacity

    If n = 0 Then
        ws.Cells(FIRST_ITEM_ROW, COL_ITEM).Value = "Nothing due or scheduled today"
        ws.Cells(FIRST_ITEM_ROW, COL_ITEM).Font.Italic = True
        ws.Cells(FIRST_ITEM_ROW, COL_ITEM).Font.Color = modDesignSystem.COLOR_INK_MUTED
    End If

    For i = 1 To shown
        r = FIRST_ITEM_ROW + i - 1
        Dim tText As String
        tText = TimeText(tm(i))
        If tmEnd(i) >= 0 Then tText = tText & ChrW(8211) & TimeText(tmEnd(i))
        ws.Cells(r, COL_TIME).Value = "'" & tText
        ws.Cells(r, COL_TIME).HorizontalAlignment = xlCenter
        ws.Cells(r, COL_ITEM).Value = itemText(i)
        ws.Cells(r, COL_TYPE).Value = kind(i)
        ws.Cells(r, COL_TYPE).HorizontalAlignment = xlCenter
        If kind(i) = "Meeting" Then
            ws.Cells(r, COL_TYPE).Font.Color = modDesignSystem.COLOR_CATEGORICAL_1
            ws.Cells(r, COL_TYPE).Font.Bold = True
        End If
        ' Already over: a meeting whose end (or start, if no end) has
        ' passed reads muted; tasks stay full-strength until done.
        Dim overAt As Double
        If tmEnd(i) >= 0 Then overAt = tmEnd(i) Else overAt = tm(i)
        If kind(i) = "Meeting" And overAt < nowFrac Then
            ws.Range(ws.Cells(r, COL_TIME), ws.Cells(r, COL_TYPE)).Font.Color = modDesignSystem.COLOR_INK_MUTED
        End If
    Next i

    Dim footer As String
    If n > shown Then footer = "+" & (n - shown) & " more"
    If overdue > 0 Then
        If Len(footer) > 0 Then footer = footer & "   " & ChrW(183) & "   "
        footer = footer & overdue & " overdue " & ChrW(8211) & " see Tasks"
    End If
    With ws.Cells(FOOTER_ROW, COL_TIME)
        .Value = footer
        .HorizontalAlignment = xlCenter
        .Font.Bold = (overdue > 0)
        .Font.Color = IIf(overdue > 0, modDesignSystem.COLOR_STATUS_CRITICAL, modDesignSystem.COLOR_INK_MUTED)
    End With

    RenderTodayPanel = "OK " & n & " " & overdue
    Exit Function
ErrHandler:
    RenderTodayPanel = "ERROR " & Err.Number & ": " & Err.Description
End Function
