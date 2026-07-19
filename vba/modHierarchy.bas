Attribute VB_Name = "modHierarchy"
Option Explicit

' Shared hierarchy-engine core (PLAN.md §6.2 points 2/3/6, §6.3). Reused by
' Tasks and Recurring Tasks -- each sheet's own thin wrapper builds a
' HierarchyColumnMap describing its own column layout and calls
' RenumberHierarchy. This module has no sheet-specific knowledge.
'
' Strategy: read every data row's full contents into memory, rebuild the
' tree from ParentKey, sort each parent's children by SiblingOrder, walk
' it depth-first to assign fresh ID/Level, then clear and rewrite the
' whole range in the new order. Repeated in-place Rows.Cut/Insert would be
' slower and more fragile than a single read-reorder-rewrite pass, and
' this already has to run on every structural change regardless.

Public Type HierarchyColumnMap
    FirstDataRow As Long
    LastCol As Long            ' rightmost column of the row range to move (the whole row block)
    ColID As Long
    ColTask As Long
    ColNumSubtasks As Long     ' see RenumberHierarchy's clear/rewrite loop for why this is excluded
    ColDaysLeft As Long        ' same reasoning as ColNumSubtasks -- 0 (VBA's default) means "not applicable"
                               ' for a sheet with no such column (e.g. Recurring Tasks), since real columns
                               ' are always >= 1 and can never accidentally match
    ColSubtaskOf As Long
    ColKey As Long
    ColParentKey As Long
    ColLevel As Long
    ColSiblingOrder As Long
End Type

Private Type HierarchyNode
    Key As Double
    EffectiveParentKey As Double   ' 0 = root; a dangling ParentKey reference is also treated as root
    RawParentText As String        ' ParentKey cell's raw text, captured in Pass 1 -- Pass 2 no longer
                                    ' re-reads the sheet since a real node's row number isn't known
                                    ' until after blank rows have been skipped (see Pass 1 below)
    SiblingOrder As Double
    NewID As String
    NewLevel As Long
    WasCollapsed As Boolean
    Values() As Variant            ' full row snapshot, columns 1..LastCol
End Type

' .End(xlUp) is unreliable for finding a hierarchy sheet's last data row:
' when the TRUE last row is hidden (e.g. a collapsed subtask block
' extending to the bottom of the current data), it silently returns the
' last VISIBLE row instead, even though the hidden row's data is fully
' intact -- confirmed via live COM testing (a collapsed Grandchild row's
' Task/TaskKey read back correctly via direct cell access, but
' End(xlUp).Row reported the row above it). Every hierarchy-engine
' function that needs "the last row with a key" uses this instead of
' End(xlUp) directly. 5000 is a generous fixed scan bound, comfortably
' beyond any hierarchy sheet's own row allowance (Tasks: 500, Recurring
' Tasks: 100) -- a full scan of that many cells is still trivially fast.
Public Function LastHierarchyRow(ws As Worksheet, keyCol As Long, firstDataRow As Long) As Long
    Dim r As Long
    Dim result As Long
    result = firstDataRow - 1
    For r = firstDataRow To 5000
        If Len(Trim$(ws.Cells(r, keyCol).Value & "")) > 0 Then
            result = r
        End If
    Next r
    LastHierarchyRow = result
End Function

Public Sub RenumberHierarchy(ws As Worksheet, colMap As HierarchyColumnMap)
    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, colMap.ColKey, colMap.FirstDataRow)
    If lastRow < colMap.FirstDataRow Then Exit Sub   ' no data rows

    Dim collapsedKeys As Object
    Set collapsedKeys = CreateObject("Scripting.Dictionary")

    Dim indexByKey As Object
    Set indexByKey = CreateObject("Scripting.Dictionary")

    Dim i As Long, r As Long, c As Long

    ' Pass 1: snapshot every row with a real (non-blank) Key -- skipping
    ' rows entirely if their Key cell is blank, rather than snapshotting
    ' them as a phantom node. A truly blank row's Key would otherwise read
    ' back as CDbl(Empty) = 0, colliding with the 0 sentinel WalkChildren
    ' uses to mean "root" -- confirmed via live COM testing to send
    ' WalkChildren's recursive call right back into the *same* root
    ' children group (since a phantom node's own Key is also 0), growing
    ' cursor past rowCount until orderedIndexes(cursor) throws Subscript
    ' out of range. Rows: 500-row (Tasks) / 100-row (Recurring Tasks)
    ' pre-formatted tables mean a user can plausibly type into a non-
    ' contiguous blank row, leaving genuinely blank rows inside the scan
    ' range -- not just a theoretical edge case. rowCount is now the
    ' count of REAL nodes only; the final rewrite loop below already
    ' writes starting at FirstDataRow, so this also naturally compacts
    ' away the gap instead of just tolerating it.
    Dim rowCount As Long
    rowCount = 0
    For r = colMap.FirstDataRow To lastRow
        If Len(Trim$(ws.Cells(r, colMap.ColKey).Value & "")) > 0 Then rowCount = rowCount + 1
    Next r

    Dim nodes() As HierarchyNode
    ReDim nodes(1 To rowCount)

    i = 0
    For r = colMap.FirstDataRow To lastRow
        If Len(Trim$(ws.Cells(r, colMap.ColKey).Value & "")) > 0 Then
            i = i + 1
            With nodes(i)
                .Key = CDbl(ws.Cells(r, colMap.ColKey).Value)
                .SiblingOrder = CDbl(ws.Cells(r, colMap.ColSiblingOrder).Value)
                .RawParentText = Trim$(ws.Cells(r, colMap.ColParentKey).Value & "")
                .WasCollapsed = ws.Rows(r).Hidden
                If .WasCollapsed Then collapsedKeys(.Key) = True

                Dim vals() As Variant
                ReDim vals(1 To colMap.LastCol)
                For c = 1 To colMap.LastCol
                    vals(c) = ws.Cells(r, c).Value
                Next c
                .Values = vals
            End With
            indexByKey(nodes(i).Key) = i
        End If
    Next r

    ' Pass 2: resolve EffectiveParentKey now that indexByKey is complete --
    ' blank ParentKey and a dangling (non-existent) ParentKey both resolve
    ' to 0 (root), so no row can ever be silently dropped from the walk
    ' below even if a ParentKey was left pointing at a since-deleted task.
    For i = 1 To rowCount
        If Len(nodes(i).RawParentText) = 0 Then
            nodes(i).EffectiveParentKey = 0
        Else
            Dim pk As Double
            pk = CDbl(nodes(i).RawParentText)
            If indexByKey.Exists(pk) Then
                nodes(i).EffectiveParentKey = pk
            Else
                nodes(i).EffectiveParentKey = 0   ' dangling reference -> treat as root
            End If
        End If
    Next i

    ' Group children by effective parent key (0 = root) and sort each
    ' group by SiblingOrder.
    Dim childrenOf As Object
    Set childrenOf = CreateObject("Scripting.Dictionary")
    For i = 1 To rowCount
        Dim key As Double
        key = nodes(i).EffectiveParentKey
        If Not childrenOf.Exists(key) Then
            Dim newColl As Collection
            Set newColl = New Collection
            childrenOf.Add key, newColl
        End If
        childrenOf(key).Add i
    Next i

    Dim k As Variant
    For Each k In childrenOf.Keys
        SortChildrenBySiblingOrder childrenOf(k), nodes
    Next k

    ' Depth-first walk from root (key 0), assigning NewID / NewLevel and
    ' producing the new row order. Every node is reachable from 0 by
    ' construction (Pass 2 above), so orderedIndexes always ends up with
    ' exactly rowCount entries -- no separate orphan-recovery pass needed.
    Dim orderedIndexes() As Long
    ReDim orderedIndexes(1 To rowCount)
    Dim cursor As Long
    cursor = 0
    WalkChildren childrenOf, nodes, 0, "", 0, orderedIndexes, cursor
    Debug.Assert cursor = rowCount

    ' Clear the existing data range and rewrite in new order -- EXCEPT
    ' ColNumSubtasks and ColDaysLeft. Like SortKey/CampaignSortHelper/etc
    ' (see this module's header comment), Number of Subtasks
    ' (hierarchy.py's number_of_subtasks_formula) and Tasks' own Days Left
    ' (tasks.py: "=IF(Deadline="","",Deadline-TODAY())") are each a
    ' per-row formula referencing only that same row's own cells, so both
    ' are self-healing as data moves underneath them -- clearing and
    ' rewriting either as a captured *value* would instead freeze whatever
    ' it happened to compute at snapshot time, destroying the live formula
    ' on every single renumber. Confirmed via live testing (Number of
    ' Subtasks read back as a stale 0 after a second RenumberHierarchy
    ' call; Days Left found blank after Phase 5's own manual click-through,
    ' the exact same bug class, just never given the same exclusion when
    ' it was first fixed) -- both bugs predate their exclusion and affect
    ' any sheet built on this engine (Tasks and Recurring Tasks alike),
    ' not something new.
    Dim clearRange As Range
    For c = 1 To colMap.LastCol
        If c <> colMap.ColNumSubtasks And c <> colMap.ColDaysLeft Then
            Dim oneCol As Range
            Set oneCol = ws.Range(ws.Cells(colMap.FirstDataRow, c), ws.Cells(lastRow, c))
            If clearRange Is Nothing Then
                Set clearRange = oneCol
            Else
                Set clearRange = Union(clearRange, oneCol)
            End If
        End If
    Next c
    clearRange.ClearContents
    ws.Rows(colMap.FirstDataRow & ":" & lastRow).Hidden = False

    ' Force the ID and Subtask-of columns to Text format BEFORE writing --
    ' otherwise Excel silently reinterprets a numeric-looking ID string
    ' ("1", "1.1", "1.2") as an actual number. That's not just cosmetic:
    ' ID "1.10" and ID "1.1" would collide as the identical number (1.1)
    ' for any task with a 10th+ subtask under the same parent, and the
    ' COUNTIF wildcard match in Number of Subtasks (PLAN.md §6.2 point 3)
    ' requires ID to be text. Three-part IDs like "1.2.1" happen to
    ' survive as text on their own since they're not valid numbers, which
    ' is what let this bug hide during initial testing.
    ws.Range(ws.Cells(colMap.FirstDataRow, colMap.ColID), ws.Cells(lastRow, colMap.ColID)).NumberFormat = "@"
    ws.Range(ws.Cells(colMap.FirstDataRow, colMap.ColSubtaskOf), ws.Cells(lastRow, colMap.ColSubtaskOf)).NumberFormat = "@"

    For i = 1 To rowCount
        r = colMap.FirstDataRow + i - 1
        Dim srcIdx As Long
        srcIdx = orderedIndexes(i)

        For c = 1 To colMap.LastCol
            If c <> colMap.ColNumSubtasks And c <> colMap.ColDaysLeft Then
                ws.Cells(r, c).Value = nodes(srcIdx).Values(c)
            End If
        Next c
        ws.Cells(r, colMap.ColID).Value = nodes(srcIdx).NewID
        ws.Cells(r, colMap.ColLevel).Value = nodes(srcIdx).NewLevel
        ws.Cells(r, colMap.ColTask).IndentLevel = nodes(srcIdx).NewLevel

        If nodes(srcIdx).EffectiveParentKey <> 0 Then
            Dim parentIdx As Long
            parentIdx = indexByKey(nodes(srcIdx).EffectiveParentKey)
            ws.Cells(r, colMap.ColSubtaskOf).Value = nodes(parentIdx).NewID
        Else
            ws.Cells(r, colMap.ColSubtaskOf).Value = ""
        End If

        If collapsedKeys.Exists(nodes(srcIdx).Key) Then
            ws.Rows(r).Hidden = True
        End If
    Next i
End Sub

Private Sub WalkChildren(childrenOf As Object, nodes() As HierarchyNode, parentKey As Double, _
                         parentID As String, level As Long, orderedIndexes() As Long, ByRef cursor As Long)
    If Not childrenOf.Exists(parentKey) Then Exit Sub
    Dim siblings As Collection
    Set siblings = childrenOf(parentKey)

    Dim position As Long
    position = 0
    Dim idx As Variant
    For Each idx In siblings
        position = position + 1
        cursor = cursor + 1
        orderedIndexes(cursor) = idx

        Dim newID As String
        If Len(parentID) = 0 Then
            newID = CStr(position)
        Else
            newID = parentID & "." & CStr(position)
        End If
        nodes(idx).NewID = newID
        nodes(idx).NewLevel = level

        WalkChildren childrenOf, nodes, nodes(idx).Key, newID, level + 1, orderedIndexes, cursor
    Next idx
End Sub

Private Sub SortChildrenBySiblingOrder(siblings As Collection, nodes() As HierarchyNode)
    ' Insertion sort on a Collection of node indexes, by SiblingOrder --
    ' sibling counts are small (a task's direct children), no need for
    ' anything fancier.
    Dim n As Long
    n = siblings.Count
    If n < 2 Then Exit Sub

    Dim arr() As Long
    ReDim arr(1 To n)
    Dim i As Long
    For i = 1 To n
        arr(i) = siblings(i)
    Next i

    Dim j As Long, tmp As Long
    For i = 2 To n
        tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If nodes(arr(j)).SiblingOrder > nodes(tmp).SiblingOrder Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = tmp
    Next i

    Do While siblings.Count > 0
        siblings.Remove 1
    Loop
    For i = 1 To n
        siblings.Add arr(i)
    Next i
End Sub

Public Function NextSiblingOrder(ws As Worksheet, colMap As HierarchyColumnMap, parentKeyValue As Variant) As Double
    ' PLAN.md §6.2 point 6: appending a task at the end of its siblings
    ' assigns max(existing siblings' SiblingOrder) + 10. Pass an empty
    ' string/0 for parentKeyValue to compute the next top-level root order.
    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, colMap.ColKey, colMap.FirstDataRow)
    If lastRow < colMap.FirstDataRow Then
        NextSiblingOrder = 10
        Exit Function
    End If

    Dim targetParent As String
    targetParent = Trim$(parentKeyValue & "")

    Dim maxOrder As Double
    Dim found As Boolean
    Dim r As Long
    For r = colMap.FirstDataRow To lastRow
        ' Skip a genuinely blank row (no TaskKey) -- its blank ParentKey
        ' would otherwise match targetParent="" and get counted as a root-
        ' level sibling even though it isn't a real task (see
        ' RenumberHierarchy's matching Pass 1 fix for the full rationale).
        If Len(Trim$(ws.Cells(r, colMap.ColKey).Value & "")) > 0 Then
            Dim rowParent As String
            rowParent = Trim$(ws.Cells(r, colMap.ColParentKey).Value & "")
            If rowParent = targetParent Then
                found = True
                If ws.Cells(r, colMap.ColSiblingOrder).Value > maxOrder Then
                    maxOrder = ws.Cells(r, colMap.ColSiblingOrder).Value
                End If
            End If
        End If
    Next r

    If found Then
        NextSiblingOrder = maxOrder + 10
    Else
        NextSiblingOrder = 10
    End If
End Function
