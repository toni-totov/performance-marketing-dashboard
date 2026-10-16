Attribute VB_Name = "modHierarchy"
Option Explicit

' Shared hierarchy-engine core (PLAN.md §6.2 points 2/3/6, §6.3). Reused by
' Tasks and Recurring Tasks -- each sheet's own thin wrapper builds a
' HierarchyColumnMap describing its own column layout and calls
' RenumberHierarchy. This module has no sheet-specific knowledge.
'
' Strategy: read every data row's full contents into memory, rebuild the
' tree from ParentKey, sort each parent's children by SiblingOrder (and by
' PIC first, when ColPIC > 0 -- see SortChildrenBySiblingOrder), walk it
' depth-first to assign fresh ID/Level, then clear and rewrite the whole
' range in the new order. Repeated in-place Rows.Cut/Insert would be
' slower and more fragile than a single read-reorder-rewrite pass, and
' this already has to run on every structural change regardless.
'
' Performance: every read and write of the data range goes through a
' single bulk Range.Value array transfer rather than per-cell .Cells(r,c)
' COM calls -- a plain range read/write of N rows x M columns as
' individual cell accesses costs N*M COM round-trips, each with enough
' fixed per-call overhead that a 1000+ row sheet became visibly slow on
' every single structural change (add/complete/delete/reorder). Bulk
' Range.Value collapses that to a small constant number of COM calls
' regardless of sheet size. The two exceptions are Worksheet.Rows(r).Hidden
' and Range.IndentLevel, neither of which is a plain cell value reachable
' via Range.Value -- those remain per-row calls, but only for rows that
' actually need collapsing/indenting, not every row unconditionally.

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
    ColPIC As Long             ' 0 = no PIC-based grouping (default -- every existing caller leaves this
                               ' unset, so SortChildrenBySiblingOrder's PIC branch never activates for them)
    KeepFromCol As Long        ' Optional contiguous block of self-healing formula columns INSIDE 1..LastCol
    KeepToCol As Long          ' that must never be cleared/rewritten (same treatment as ColNumSubtasks/
                               ' ColDaysLeft). 0 = none. Lets a sheet carry a moved value column to the
                               ' right of its formula helpers (Tasks' RecurringSourceKey, PLAN.md §22.4).
End Type

Private Type HierarchyNode
    Key As Double
    EffectiveParentKey As Double   ' 0 = root; a dangling ParentKey reference is also treated as root
    RawParentText As String        ' ParentKey cell's raw text, captured in Pass 1 -- Pass 2 no longer
                                    ' re-reads the sheet since a real node's row number isn't known
                                    ' until after blank rows have been skipped (see Pass 1 below)
    SiblingOrder As Double
    PIC As String                  ' Trim$ of the PIC cell; "" when ColPIC=0 or the cell is blank
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
' End(xlUp) directly. 20000 is a generous fixed scan bound, comfortably
' beyond any hierarchy sheet's own row allowance (Recurring Tasks: 100;
' Tasks grows in +500 increments via modTasks.bas's row extender -- see
' its own design note -- so this bound needs to stay well ahead of that
' growth, not just the allowance at any one point in time). A single bulk
' read of the whole scan column (below) keeps this fast regardless of the
' 20000-row bound -- one COM call, not up to 20000 individual ones.
Public Function LastHierarchyRow(ws As Worksheet, keyCol As Long, firstDataRow As Long) As Long
    Dim result As Long
    result = firstDataRow - 1

    Dim vals As Variant
    vals = ws.Range(ws.Cells(firstDataRow, keyCol), ws.Cells(20000, keyCol)).Value

    Dim r As Long
    For r = 1 To UBound(vals, 1)
        If Len(Trim$(vals(r, 1) & "")) > 0 Then
            result = firstDataRow + r - 1
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

    ' Single bulk read of the entire data block -- every subsequent access
    ' in Pass 1 indexes into this in-memory array instead of making its own
    ' COM call. block(n, c) corresponds to sheet row (FirstDataRow + n - 1),
    ' column c.
    Dim block As Variant
    block = ws.Range(ws.Cells(colMap.FirstDataRow, 1), ws.Cells(lastRow, colMap.LastCol)).Value

    ' Pass 1: snapshot every row with a real (non-blank) Key -- skipping
    ' rows entirely if their Key cell is blank, rather than snapshotting
    ' them as a phantom node. A truly blank row's Key would otherwise read
    ' back as CDbl(Empty) = 0, colliding with the 0 sentinel WalkChildren
    ' uses to mean "root" -- confirmed via live COM testing to send
    ' WalkChildren's recursive call right back into the *same* root
    ' children group (since a phantom node's own Key is also 0), growing
    ' cursor past rowCount until orderedIndexes(cursor) throws Subscript
    ' out of range. Rows: 1500-row (Tasks) / 100-row (Recurring Tasks)
    ' pre-formatted tables mean a user can plausibly type into a non-
    ' contiguous blank row, leaving genuinely blank rows inside the scan
    ' range -- not just a theoretical edge case. rowCount is now the
    ' count of REAL nodes only; the final rewrite loop below already
    ' writes starting at FirstDataRow, so this also naturally compacts
    ' away the gap instead of just tolerating it.
    Dim rowCount As Long
    rowCount = 0
    For r = colMap.FirstDataRow To lastRow
        If Len(Trim$(block(r - colMap.FirstDataRow + 1, colMap.ColKey) & "")) > 0 Then rowCount = rowCount + 1
    Next r

    Dim nodes() As HierarchyNode
    ReDim nodes(1 To rowCount)

    i = 0
    For r = colMap.FirstDataRow To lastRow
        Dim blockRow As Long
        blockRow = r - colMap.FirstDataRow + 1
        If Len(Trim$(block(blockRow, colMap.ColKey) & "")) > 0 Then
            i = i + 1
            With nodes(i)
                .Key = CDbl(block(blockRow, colMap.ColKey))
                .SiblingOrder = CDbl(block(blockRow, colMap.ColSiblingOrder))
                .RawParentText = Trim$(block(blockRow, colMap.ColParentKey) & "")
                .WasCollapsed = ws.Rows(r).Hidden   ' not a plain cell value -- stays a per-row COM call
                If .WasCollapsed Then collapsedKeys(.Key) = True

                If colMap.ColPIC > 0 Then
                    .PIC = Trim$(block(blockRow, colMap.ColPIC) & "")
                End If

                Dim vals() As Variant
                ReDim vals(1 To colMap.LastCol)
                For c = 1 To colMap.LastCol
                    vals(c) = block(blockRow, c)
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
    ' group by PIC (when active) then SiblingOrder.
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

    Dim usePIC As Boolean
    usePIC = (colMap.ColPIC > 0)

    Dim k As Variant
    For Each k In childrenOf.Keys
        SortChildrenBySiblingOrder childrenOf(k), nodes, usePIC
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
    ' Debug.Assert deliberately NOT used here -- confirmed live (found
    ' while debugging a new caller, PLAN.md manager-variant work) that a
    ' FAILING Debug.Assert under COM automation (no interactive VBE
    ' session to resume a break) manifests as Application.Run hanging
    ' indefinitely, with no error and no exception -- the exact
    ' "invalid named argument" symptom CLAUDE.md already documents for a
    ' different cause, but this is a second, distinct way to trigger the
    ' identical-looking hang. A real Err.Raise is always catchable by an
    ' On Error handler regardless of automation context, so any future
    ' violation of this invariant surfaces as a normal diagnosable error
    ' instead of a silent hang that looks like a totally different bug.
    If cursor <> rowCount Then
        Err.Raise vbObjectError + 1001, "RenumberHierarchy", _
            "Tree walk reached " & cursor & " of " & rowCount & " nodes -- some row's " & _
            "ParentKey chain doesn't lead back to root. Check for a cycle or a key collision."
    End If

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
        If c <> colMap.ColNumSubtasks And c <> colMap.ColDaysLeft And Not (c >= colMap.KeepFromCol And c <= colMap.KeepToCol And colMap.KeepFromCol > 0) Then
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

    ' Build the full output block in memory first (pure VBA array writes,
    ' no COM), then flush it to the sheet as a small number of bulk
    ' Range.Value writes -- see the segmented-write loop below for why
    ' ColNumSubtasks/ColDaysLeft can't just be included in one single
    ' full-width write.
    Dim excluded() As Boolean
    ReDim excluded(1 To colMap.LastCol)
    If colMap.ColNumSubtasks > 0 And colMap.ColNumSubtasks <= colMap.LastCol Then excluded(colMap.ColNumSubtasks) = True
    If colMap.ColDaysLeft > 0 And colMap.ColDaysLeft <= colMap.LastCol Then excluded(colMap.ColDaysLeft) = True
    If colMap.KeepFromCol > 0 Then
        For c = colMap.KeepFromCol To colMap.KeepToCol
            If c <= colMap.LastCol Then excluded(c) = True
        Next c
    End If

    Dim outBlock() As Variant
    ReDim outBlock(1 To rowCount, 1 To colMap.LastCol)

    Dim srcIdx As Long, parentIdx As Long
    For i = 1 To rowCount
        srcIdx = orderedIndexes(i)
        For c = 1 To colMap.LastCol
            If Not excluded(c) Then
                outBlock(i, c) = nodes(srcIdx).Values(c)
            End If
        Next c
        outBlock(i, colMap.ColID) = nodes(srcIdx).NewID
        outBlock(i, colMap.ColLevel) = nodes(srcIdx).NewLevel

        If nodes(srcIdx).EffectiveParentKey <> 0 Then
            parentIdx = indexByKey(nodes(srcIdx).EffectiveParentKey)
            outBlock(i, colMap.ColSubtaskOf) = nodes(parentIdx).NewID
        Else
            outBlock(i, colMap.ColSubtaskOf) = ""
        End If
    Next i

    ' Flush outBlock to the sheet in contiguous column runs, skipping the
    ' excluded (self-healing formula) columns entirely -- their cells were
    ' never cleared above and must stay completely untouched, not
    ' overwritten with a stale Pass-1 snapshot value.
    Dim runStart As Long
    runStart = 0
    For c = 1 To colMap.LastCol + 1
        Dim stopRun As Boolean
        stopRun = (c > colMap.LastCol)
        If Not stopRun Then stopRun = excluded(c)

        If stopRun Then
            If runStart > 0 Then
                Dim runEnd As Long
                runEnd = c - 1
                Dim subBlock() As Variant
                ReDim subBlock(1 To rowCount, 1 To (runEnd - runStart + 1))
                Dim rr As Long, cc As Long
                For rr = 1 To rowCount
                    For cc = runStart To runEnd
                        subBlock(rr, cc - runStart + 1) = outBlock(rr, cc)
                    Next cc
                Next rr
                ws.Range(ws.Cells(colMap.FirstDataRow, runStart), ws.Cells(colMap.FirstDataRow + rowCount - 1, runEnd)).Value = subBlock
            End If
            runStart = 0
        Else
            If runStart = 0 Then runStart = c
        End If
    Next c

    ' IndentLevel and Rows.Hidden aren't plain cell values reachable via
    ' Range.Value -- these stay per-row COM calls, but only for rows that
    ' actually need indenting/hiding.
    For i = 1 To rowCount
        r = colMap.FirstDataRow + i - 1
        srcIdx = orderedIndexes(i)
        ws.Cells(r, colMap.ColTask).IndentLevel = nodes(srcIdx).NewLevel
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

' Plain single-key insertion sort on SiblingOrder when usePIC is False
' (every existing caller -- Recurring Tasks, modMigration.bas, employee
' Tasks -- takes this branch, byte-identical to this function's original
' behavior). When usePIC is True (Tasks Manager only, PLAN.md's manager-
' variant design), the sort becomes two-key: PIC first (case-insensitive
' text compare, blank sorts AFTER every named cluster -- see
' ComparePICForSort below), SiblingOrder as tiebreak within a PIC.
' > **Corrected after live use:** the original design put blank-PIC tasks
' > FIRST (ahead of every named cluster), reasoning that they'd "stay
' > ungrouped." In practice this meant a brand-new task (which always
' > starts unassigned) visually jumped to the TOP of the list, which the
' > user reported as surprising/unwanted after actually using the manager
' > build -- they want a new task to stay at the bottom until a PIC is
' > assigned. Flipped to sort blank PIC last instead.
Private Sub SortChildrenBySiblingOrder(siblings As Collection, nodes() As HierarchyNode, ByVal usePIC As Boolean)
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
            Dim shouldShift As Boolean
            If usePIC Then
                Dim picCmp As Long
                picCmp = ComparePICForSort(nodes(arr(j)).PIC, nodes(tmp).PIC)
                shouldShift = (picCmp > 0) Or (picCmp = 0 And nodes(arr(j)).SiblingOrder > nodes(tmp).SiblingOrder)
            Else
                shouldShift = (nodes(arr(j)).SiblingOrder > nodes(tmp).SiblingOrder)
            End If

            If shouldShift Then
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

' Ascending-sort comparator for the PIC key: blank always sorts AFTER any
' named PIC (both blanks compare equal), non-blank-vs-non-blank falls
' through to a plain case-insensitive text compare. Returns <0 if picA
' sorts before picB, >0 if after, 0 if equal -- same contract as StrComp,
' which this replaces specifically for SortChildrenBySiblingOrder's usePIC
' branch (StrComp alone would put "" before any real name, the opposite of
' what's wanted here -- see that Sub's own correction note).
Private Function ComparePICForSort(picA As String, picB As String) As Long
    Dim aBlank As Boolean, bBlank As Boolean
    aBlank = (Len(picA) = 0)
    bBlank = (Len(picB) = 0)

    If aBlank And bBlank Then
        ComparePICForSort = 0
    ElseIf aBlank Then
        ComparePICForSort = 1
    ElseIf bBlank Then
        ComparePICForSort = -1
    Else
        ComparePICForSort = StrComp(picA, picB, vbTextCompare)
    End If
End Function

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

' Same as NextSiblingOrder, but scoped by the tuple (ParentKey, PIC)
' instead of just ParentKey -- so a newly-(re)PIC'd task lands at the end
' of its new PIC cluster rather than wherever its old SiblingOrder happens
' to sort it. Only meaningful when colMap.ColPIC > 0 (Tasks Manager);
' passing a blank picValue computes the next order within the
' "unassigned" group for that parent, consistent with how blank PIC is
' treated as its own ungrouped-but-still-ordered set in
' SortChildrenBySiblingOrder above.
Public Function NextSiblingOrderForPIC(ws As Worksheet, colMap As HierarchyColumnMap, parentKeyValue As Variant, picValue As Variant) As Double
    Dim lastRow As Long
    lastRow = LastHierarchyRow(ws, colMap.ColKey, colMap.FirstDataRow)
    If lastRow < colMap.FirstDataRow Then
        NextSiblingOrderForPIC = 10
        Exit Function
    End If

    Dim targetParent As String
    targetParent = Trim$(parentKeyValue & "")
    Dim targetPIC As String
    targetPIC = Trim$(picValue & "")

    Dim maxOrder As Double
    Dim found As Boolean
    Dim r As Long
    For r = colMap.FirstDataRow To lastRow
        If Len(Trim$(ws.Cells(r, colMap.ColKey).Value & "")) > 0 Then
            Dim rowParent As String
            rowParent = Trim$(ws.Cells(r, colMap.ColParentKey).Value & "")
            If rowParent = targetParent Then
                Dim rowPIC As String
                rowPIC = Trim$(ws.Cells(r, colMap.ColPIC).Value & "")
                If StrComp(rowPIC, targetPIC, vbTextCompare) = 0 Then
                    found = True
                    If ws.Cells(r, colMap.ColSiblingOrder).Value > maxOrder Then
                        maxOrder = ws.Cells(r, colMap.ColSiblingOrder).Value
                    End If
                End If
            End If
        End If
    Next r

    If found Then
        NextSiblingOrderForPIC = maxOrder + 10
    Else
        NextSiblingOrderForPIC = 10
    End If
End Function
