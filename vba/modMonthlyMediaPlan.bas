Attribute VB_Name = "modMonthlyMediaPlan"
Option Explicit

' Monthly Media Plan sheet business logic (PLAN.md §9.5). The exchange-
' rate popup itself needs no wrapper here -- Monthly Media Plan.bas calls
' frmExchangeRate.ShowRate directly, same as Tasks/Archive/Campaigns call
' frmNotes.ShowForCell directly for Notes.

Private Const COL_FUNDED As Long = 1
Private Const COL_CAMPAIGN_NAME As Long = 2
Private Const COL_PLATFORM As Long = 3
Private Const COL_OPTIMIZATION As Long = 4
Private Const COL_AUDIENCE As Long = 5
Private Const COL_START_DATE As Long = 6
Private Const COL_END_DATE As Long = 7
Private Const COL_BUDGET_EUR As Long = 8
' Budget (USD), column 9, is a self-healing formula (IF(BudgetEUR="","",
' BudgetEUR*MonthlyExchangeRate)) -- never cleared, same reasoning as
' Tasks' SortKey/Number of Subtasks (PLAN.md §19 Phase 1 correction note).

Private Const FIRST_DATA_ROW As Long = 7
Private Const LAST_DATA_ROW As Long = 36   ' the fixed 30-row allowance, sheets/monthly_media_plan.py

Private Const HIDDEN_DATE_ROW As Long = 5
Private Const GANTT_FIRST_COL As Long = 10   ' J
Private Const NUM_GANTT_WEEKS As Long = 6

Private Function MMPSheet() As Worksheet
    Set MMPSheet = ThisWorkbook.Sheets("Monthly Media Plan")
End Function

' Hides any of the 6 pre-built Gantt week columns that don't overlap
' PlanYear/PlanMonth at all (§19 Phase 5, user-requested/found live) --
' VBA-side equivalent of sheets/monthly_media_plan.py's
' _needed_gantt_weeks(), needed here because PlanYear/PlanMonth can
' change at runtime via "Start New Month" (ApplyMonthlyReset), unlike
' the Python function which only runs once at initial build. See that
' function's own comment for why the Gantt always builds 6 columns but
' most months only need 4 or 5 of them visible. ws.Calculate first,
' forcing the week-start formulas in HIDDEN_DATE_ROW to reflect
' PlanYear/PlanMonth's just-changed value before reading them.
Public Sub UpdateGanttWeekVisibility()
    Dim ws As Worksheet
    Set ws = MMPSheet()
    ws.Calculate

    Dim planYear As Long, planMonth As Long
    planYear = ws.Range("PlanYear").Value
    planMonth = ws.Range("PlanMonth").Value

    Dim monthStart As Date, monthEnd As Date
    monthStart = DateSerial(planYear, planMonth, 1)
    monthEnd = DateSerial(planYear, planMonth + 1, 1) - 1

    Dim w As Long, weekStart As Date, weekEnd As Date, col As Long
    For w = 0 To NUM_GANTT_WEEKS - 1
        weekStart = ws.Cells(HIDDEN_DATE_ROW, GANTT_FIRST_COL + w).Value
        weekEnd = weekStart + 6
        col = GANTT_FIRST_COL + w
        ws.Columns(col).Hidden = Not (weekEnd >= monthStart And weekStart <= monthEnd)
    Next w
End Sub

' The actual clearing+month-advance action (§9.5), assuming the
' confirmation has already happened. Public and MsgBox-free specifically
' so it's testable directly via Application.Run, bypassing
' ResetMonthlyPlan's confirm dialog -- same reasoning as Tasks'
' CascadeDelete/WouldCreateCycle split.
Public Sub ApplyMonthlyReset()
    Dim ws As Worksheet
    Set ws = MMPSheet()

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim clearRange As Range
    Set clearRange = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_FUNDED), ws.Cells(LAST_DATA_ROW, COL_BUDGET_EUR))
    clearRange.ClearContents

    ws.Range("PlanYear").Value = Year(Date)
    ws.Range("PlanMonth").Value = Month(Date)

    UpdateGanttWeekVisibility

    Application.EnableEvents = prevEvents
End Sub

' Production entry point, wired to the "Start New Month" button (§9.5) --
' confirms via MsgBox (destructive, consistent with Tasks' own confirm-
' before-delete pattern) before calling ApplyMonthlyReset.
Public Sub ResetMonthlyPlan()
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox( _
        "This will clear all campaigns from the current Monthly Media Plan and start a new month. This cannot be undone. Continue?", _
        vbYesNo + vbExclamation, "Start New Month")
    If confirmResult = vbYes Then
        ApplyMonthlyReset
    End If
End Sub
