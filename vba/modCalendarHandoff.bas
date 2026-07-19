Attribute VB_Name = "modCalendarHandoff"
Option Explicit

' Shared routines for the generic Calendar handoff contract (PLAN.md
' §8.1). The four named ranges (HandoffOriginSheet, HandoffKey,
' HandoffKeyColumn, HandoffTargetColumn) are workbook-scoped defined names
' pointing at reserved cells on the Calendar sheet (see
' sheets/calendar_handoff.py, built during Phase 1/2). Any origin sheet's
' VBA calls SetCalendarHandoff before navigating to Calendar; Calendar's
' Confirm button (§8.4) calls ResolveCalendarHandoffTarget to know where
' to write the picked value back to, then ClearCalendarHandoff.

Public Sub SetCalendarHandoff(originSheet As String, handoffKey As Variant, keyColumn As String, targetColumn As String)
    ThisWorkbook.Names("HandoffOriginSheet").RefersToRange.Value = originSheet
    ThisWorkbook.Names("HandoffKey").RefersToRange.Value = handoffKey
    ThisWorkbook.Names("HandoffKeyColumn").RefersToRange.Value = keyColumn
    ThisWorkbook.Names("HandoffTargetColumn").RefersToRange.Value = targetColumn
End Sub

Public Sub ClearCalendarHandoff()
    ThisWorkbook.Names("HandoffOriginSheet").RefersToRange.Value = ""
    ThisWorkbook.Names("HandoffKey").RefersToRange.Value = ""
    ThisWorkbook.Names("HandoffKeyColumn").RefersToRange.Value = ""
    ThisWorkbook.Names("HandoffTargetColumn").RefersToRange.Value = ""
End Sub

Public Function HasPendingCalendarHandoff() As Boolean
    HasPendingCalendarHandoff = Len(Trim$(ThisWorkbook.Names("HandoffOriginSheet").RefersToRange.Value & "")) > 0
End Function

' Resolves the handoff to the actual target Range on the origin sheet by
' scanning for the row where HandoffKeyColumn's column equals HandoffKey --
' NOT by row number, since rows may have moved (renumbering, manual
' reorder) between when the handoff was set and when Confirm is clicked.
' Returns Nothing if no handoff is pending or the target can't be resolved
' (origin sheet/columns missing, key not found).
Public Function ResolveCalendarHandoffTarget() As Range
    If Not HasPendingCalendarHandoff() Then
        Set ResolveCalendarHandoffTarget = Nothing
        Exit Function
    End If

    Dim originSheet As String, handoffKey As String
    Dim keyColumnHeader As String, targetColumnHeader As String
    originSheet = ThisWorkbook.Names("HandoffOriginSheet").RefersToRange.Value
    handoffKey = Trim$(ThisWorkbook.Names("HandoffKey").RefersToRange.Value & "")
    keyColumnHeader = ThisWorkbook.Names("HandoffKeyColumn").RefersToRange.Value
    targetColumnHeader = ThisWorkbook.Names("HandoffTargetColumn").RefersToRange.Value

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(originSheet)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ResolveCalendarHandoffTarget = Nothing
        Exit Function
    End If

    Dim keyCol As Long, targetCol As Long
    keyCol = FindColumnByHeader(ws, keyColumnHeader)
    targetCol = FindColumnByHeader(ws, targetColumnHeader)
    If keyCol = 0 Or targetCol = 0 Then
        Set ResolveCalendarHandoffTarget = Nothing
        Exit Function
    End If

    ' modHierarchy.LastHierarchyRow, not .End(xlUp) -- the origin sheet
    ' may be a hierarchy sheet (Tasks, Recurring Tasks) where a collapsed
    ' block can leave the true last row hidden, which .End(xlUp) silently
    ' misreports as the last *visible* row (see PLAN.md §19 Phase 1's
    ' correction note). Safe for any origin sheet regardless of whether it
    ' actually hides rows.
    Dim lastRow As Long
    lastRow = modHierarchy.LastHierarchyRow(ws, keyCol, 2)

    Dim r As Long
    For r = 2 To lastRow   ' data starts at row 2; row 1 is the header
        If Trim$(ws.Cells(r, keyCol).Value & "") = handoffKey Then
            Set ResolveCalendarHandoffTarget = ws.Cells(r, targetCol)
            Exit Function
        End If
    Next r

    Set ResolveCalendarHandoffTarget = Nothing
End Function

' .End(xlToLeft) is the column-axis analog of .End(xlUp)'s hidden-row bug
' (PLAN.md §19 Phase 1 correction note): confirmed via live COM testing
' that it silently stops at the last *visible* column, never reaching a
' hidden header like Tasks' TaskKey (column 17, hidden -- .End(xlToLeft)
' from the sheet's right edge stopped at column 16, the last visible
' column "Notes"). A hierarchy sheet's key/parent-key columns are always
' hidden, so this isn't an edge case -- it's the normal shape of exactly
' the sheets this function needs to search. Manual scan to a generous
' fixed bound instead, immune to Hidden state, mirroring
' modHierarchy.LastHierarchyRow's row-axis fix.
Private Function FindColumnByHeader(ws As Worksheet, headerText As String) As Long
    Dim lastCol As Long
    lastCol = 200
    Dim c As Long
    For c = 1 To lastCol
        If ws.Cells(1, c).Value = headerText Then
            FindColumnByHeader = c
            Exit Function
        End If
    Next c
    FindColumnByHeader = 0
End Function
