Attribute VB_Name = "modCampaigns"
Option Explicit

' Campaigns sheet business logic (PLAN.md §10.2, §10.4). Every Public
' Function/Sub here is independently callable via Application.Run for
' testing. Unlike Tasks, Campaigns has NO hierarchy engine -- it's a
' fixed 20-slot layout (sheets/campaigns.py), never grows/shrinks, so
' "deleting" a completed campaign (§10.4's wording) means CLEARING that
' slot's cells in place, never a real Rows.Delete -- physically deleting
' rows would misalign every subsequent slot's hardcoded row positions and
' corrupt their Gantt/summary formulas. None of this module's paths show
' a MsgBox (Campaigns' Status dropdown has no "Delete" option the way
' Tasks does, and §10.4 describes no confirmation step), so unlike
' modTasks.bas nothing here needs a Public-for-MsgBox-bypass split.

' Up/Down reorder buttons (§19 Phase 5, user-requested) -- always
' columns A/B, ahead of everything else (sheets/campaigns.py's
' _UP_COL/_DOWN_COL).
Private Const COL_UP As Long = 1
Private Const COL_DOWN As Long = 2

Private Const COL_CAMPAIGN_NAME As Long = 3
Private Const COL_NUM_TASKS As Long = 4
Private Const COL_START_DATE As Long = 5
Private Const COL_END_DATE As Long = 6
Private Const COL_STATUS As Long = 7
Private Const COL_PROGRESS As Long = 8
Private Const COL_NOTES As Long = 9
Private Const COL_CAMPAIGN_KEY As Long = 25
Private Const COL_CAMPAIGN_NAME_LIST As Long = 26

Private Const FIRST_DATA_ROW As Long = 3
Private Const SLOT_SIZE As Long = 10  ' main row + mini header row + 8 data rows
Private Const NUM_SLOTS As Long = 20
Private Const SUMMARY_ROWS_PER_SLOT As Long = 8

Private Const ARCHIVE_COL_CAMPAIGN_KEY As Long = 8

Private Function CampaignsSheet() As Worksheet
    Set CampaignsSheet = ThisWorkbook.Sheets("Campaigns")
End Function

Private Function IsMainRow(r As Long) As Boolean
    If r < FIRST_DATA_ROW Then
        IsMainRow = False
    Else
        IsMainRow = ((r - FIRST_DATA_ROW) Mod SLOT_SIZE) = 0
    End If
End Function

' Bootstrap point for a genuinely NEW campaign, same reasoning as Tasks'
' HandleTaskNameChange (§19 Phase 3 correction note): a slot's CampaignKey
' is blank until its Campaign Name is first typed, at which point it gets
' assigned once and never again (also never re-triggers on a rename,
' since CampaignKey is only blank for a slot that's never held real data).
Public Function HandleCampaignNameChange(r As Long) As String
    On Error GoTo ErrHandler
    If Not IsMainRow(r) Then
        HandleCampaignNameChange = "OK"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = CampaignsSheet()

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    If Len(Trim$(ws.Cells(r, COL_CAMPAIGN_NAME).Value & "")) > 0 Then
        If Len(Trim$(ws.Cells(r, COL_CAMPAIGN_KEY).Value & "")) = 0 Then
            ' Brand-new campaign: assign a key, default Status to
            ' "Planned" (user-requested, same pattern as Tasks defaulting
            ' to "Not Started" -- also what the Gantt's status-coloring
            ' needs to have something to color against right away), then
            ' compact all 20 slots so it jumps to the first available
            ' slot instead of sitting wherever it happened to be typed
            ' (§19 Phase 5, user-requested -- mirrors Tasks'
            ' RenumberHierarchy compaction).
            ws.Cells(r, COL_CAMPAIGN_KEY).Value = NextCampaignKey()
            If Len(Trim$(ws.Cells(r, COL_STATUS).Value & "")) = 0 Then
                ws.Cells(r, COL_STATUS).Value = "Planned"
            End If
            CompactCampaignSlots   ' also refreshes CampaignNameList
        Else
            ' Renaming an existing campaign -- do NOT reposition it, just
            ' keep the dropdown mirror current.
            RefreshCampaignNameList
        End If
    Else
        RefreshCampaignNameList
    End If

    Application.EnableEvents = prevEvents

    HandleCampaignNameChange = "OK"
    Exit Function
ErrHandler:
    Application.EnableEvents = prevEvents
    HandleCampaignNameChange = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Rebuilds the Linked Campaign dropdown's source list (§10.2 point 2,
' sheets/campaigns.py module docstring -- design correction found live,
' PLAN.md §19 Phase 5, TWICE). First correction: classic Excel List Data
' Validation only skips a genuinely untouched blank cell in its source
' range, never a cell holding a FORMULA that merely evaluates to "" -- so
' a per-row formula mirror (the original approach) still showed a blank,
' selectable dropdown row for every one of the ~19 typically-unused
' slots. Second correction (user still saw one trailing blank row even
' after switching to a plain-value list with unused rows ClearContents'd):
' a FIXED-size source range still isn't enough -- the named range itself
' must be dynamically sized (see build()'s OFFSET/COUNTA definition) to
' exactly the real entry count, or Excel's dropdown still renders one
' blank row for whatever's left over in the range. This function's job is
' just to keep Z4:Z23 a gap-free, VALUE-only list of real campaign names
' (Z3 itself is the permanent "None" floor, never touched here -- see
' module docstring) so that dynamic range always measures the right
' height via COUNTA. Called on any Campaign Name edit (create/rename, via
' HandleCampaignNameChange) and on archiving (ArchiveCampaign, clears a
' name), plus once on Workbook_Open to backfill a workbook that already
' had campaigns typed in before this fix.
Public Function RefreshCampaignNameList() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = CampaignsSheet()

    Dim names() As String
    ReDim names(1 To NUM_SLOTS)
    Dim n As Long
    n = 0

    Dim slot As Long, r As Long
    For slot = 0 To NUM_SLOTS - 1
        r = FIRST_DATA_ROW + slot * SLOT_SIZE
        Dim nm As String
        nm = Trim$(ws.Cells(r, COL_CAMPAIGN_NAME).Value & "")
        If Len(nm) > 0 Then
            n = n + 1
            names(n) = nm
        End If
    Next slot

    Dim i As Long
    For i = 1 To n
        ws.Cells(FIRST_DATA_ROW + i, COL_CAMPAIGN_NAME_LIST).Value = names(i)
    Next i
    If n < NUM_SLOTS Then
        ws.Range(ws.Cells(FIRST_DATA_ROW + 1 + n, COL_CAMPAIGN_NAME_LIST), _
                 ws.Cells(FIRST_DATA_ROW + NUM_SLOTS, COL_CAMPAIGN_NAME_LIST)).ClearContents
    End If

    RefreshCampaignNameList = "OK"
    Exit Function
ErrHandler:
    RefreshCampaignNameList = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Compacts all 20 slots so real campaigns always occupy a contiguous
' block starting at slot 0, with no gaps (§19 Phase 5, user-requested --
' mirrors Tasks' RenumberHierarchy: typing a new task into a non-
' contiguous row makes it "jump" to right after the last existing task,
' and this is the Campaigns equivalent). Unlike Tasks there's no
' hierarchy/SiblingOrder to sort by -- campaigns simply keep their
' current relative top-to-bottom order, with any gaps squeezed out.
' Swaps/moves only the identity fields (Name/Start/End/Status/Notes/
' CampaignKey) -- Number of Tasks/Progress/ReminderDaysLeft/
' ReminderSortKey are all position-relative formulas that auto-resolve
' correctly once those fields land in their new slot, same reasoning as
' HandleCampaignReorder. Also rewrites every slot's Up/Down icons (real
' VALUES, not formulas, so they don't auto-follow the way formula columns
' do) and forcibly collapses every slot's summary block, since a
' reshuffle can leave a stale expand/collapse state that no longer
' matches whichever campaign (if any) now occupies that slot. Called
' whenever a genuinely NEW campaign is created (HandleCampaignNameChange)
' or a campaign is archived (ArchiveCampaign, which clears a slot and
' would otherwise leave a gap behind).
Public Function CompactCampaignSlots() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = CampaignsSheet()

    Dim names(1 To NUM_SLOTS) As Variant
    Dim startDates(1 To NUM_SLOTS) As Variant
    Dim endDates(1 To NUM_SLOTS) As Variant
    Dim statusVals(1 To NUM_SLOTS) As Variant
    Dim notesVals(1 To NUM_SLOTS) As Variant
    Dim keyVals(1 To NUM_SLOTS) As Variant
    Dim n As Long
    n = 0

    Dim slot As Long, r As Long
    For slot = 0 To NUM_SLOTS - 1
        r = FIRST_DATA_ROW + slot * SLOT_SIZE
        If Len(Trim$(ws.Cells(r, COL_CAMPAIGN_NAME).Value & "")) > 0 Then
            n = n + 1
            names(n) = ws.Cells(r, COL_CAMPAIGN_NAME).Value
            startDates(n) = ws.Cells(r, COL_START_DATE).Value
            endDates(n) = ws.Cells(r, COL_END_DATE).Value
            statusVals(n) = ws.Cells(r, COL_STATUS).Value
            notesVals(n) = ws.Cells(r, COL_NOTES).Value
            keyVals(n) = ws.Cells(r, COL_CAMPAIGN_KEY).Value
        End If
    Next slot

    ' Clear every slot's identity fields, Up/Down icons, and collapse its
    ' summary block -- a clean slate before rewriting the compacted data.
    For slot = 0 To NUM_SLOTS - 1
        r = FIRST_DATA_ROW + slot * SLOT_SIZE
        ws.Cells(r, COL_CAMPAIGN_NAME).Value = ""
        ws.Cells(r, COL_START_DATE).Value = ""
        ws.Cells(r, COL_END_DATE).Value = ""
        ws.Cells(r, COL_STATUS).Value = ""
        ws.Cells(r, COL_NOTES).Value = ""
        ws.Cells(r, COL_CAMPAIGN_KEY).Value = ""
        ws.Cells(r, COL_UP).Value = ""
        ws.Cells(r, COL_DOWN).Value = ""
        ws.Rows((r + 1) & ":" & (r + 1 + SUMMARY_ROWS_PER_SLOT)).Hidden = True
    Next slot

    Dim i As Long
    For i = 1 To n
        r = FIRST_DATA_ROW + (i - 1) * SLOT_SIZE
        ws.Cells(r, COL_CAMPAIGN_NAME).Value = names(i)
        ws.Cells(r, COL_START_DATE).Value = startDates(i)
        ws.Cells(r, COL_END_DATE).Value = endDates(i)
        ws.Cells(r, COL_STATUS).Value = statusVals(i)
        ws.Cells(r, COL_NOTES).Value = notesVals(i)
        ws.Cells(r, COL_CAMPAIGN_KEY).Value = keyVals(i)
        ws.Cells(r, COL_UP).Value = ICON_UP
        ws.Cells(r, COL_DOWN).Value = ICON_DOWN
        ' Button-look applied dynamically per occupied slot (§19 Phase 6,
        ' user-reported), not baked in across all NUM_SLOTS at build time
        ' -- matches Tasks'/Recurring Tasks' own fix (modTasks.bas's
        ' RefreshButtonStyling).
        modDesignSystem.ApplyButtonLook ws.Cells(r, COL_UP)
        modDesignSystem.ApplyButtonLook ws.Cells(r, COL_DOWN)
    Next i

    RefreshCampaignNameList

    CompactCampaignSlots = "OK"
    Exit Function
ErrHandler:
    CompactCampaignSlots = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Computes the next monotonic CampaignKey. Checks both Campaigns (live)
' AND Campaigns Archive (frozen) -- an archived campaign's key must never
' collide with a new one, same reasoning as Tasks' NextTaskKey.
Private Function NextCampaignKey() As Double
    Dim ws As Worksheet, archiveWs As Worksheet
    Set ws = CampaignsSheet()
    Set archiveWs = ThisWorkbook.Sheets("Campaigns Archive")

    Dim maxKey As Double
    maxKey = 0

    Dim slot As Long, r As Long
    For slot = 0 To NUM_SLOTS - 1
        r = FIRST_DATA_ROW + slot * SLOT_SIZE
        If Len(Trim$(ws.Cells(r, COL_CAMPAIGN_KEY).Value & "")) > 0 Then
            If CDbl(ws.Cells(r, COL_CAMPAIGN_KEY).Value) > maxKey Then maxKey = CDbl(ws.Cells(r, COL_CAMPAIGN_KEY).Value)
        End If
    Next slot

    Dim archLastRow As Long
    archLastRow = modHierarchy.LastHierarchyRow(archiveWs, ARCHIVE_COL_CAMPAIGN_KEY, 2)
    For r = 2 To archLastRow
        If Len(Trim$(archiveWs.Cells(r, ARCHIVE_COL_CAMPAIGN_KEY).Value & "")) > 0 Then
            If CDbl(archiveWs.Cells(r, ARCHIVE_COL_CAMPAIGN_KEY).Value) > maxKey Then maxKey = CDbl(archiveWs.Cells(r, ARCHIVE_COL_CAMPAIGN_KEY).Value)
        End If
    Next r

    NextCampaignKey = maxKey + 1
End Function

' Reorders a campaign (§19 Phase 5, user-requested): unlike Tasks'
' SiblingOrder-based reorder, Campaigns has no hierarchy engine -- it's a
' fixed 20-slot layout (module header comment), so "moving" a campaign
' means swapping its identity fields (Name/Start/End/Status/Notes/
' CampaignKey) with the nearest slot that actually has a campaign in the
' given direction, never touching Number of Tasks/Progress/
' ReminderDaysLeft/ReminderSortKey (all position-relative formulas that
' auto-resolve correctly once the swapped fields land in their new slot)
' or the hidden summary rows (their formulas key off their OWN slot's
' CampaignKey cell, so they automatically show the newly-arrived
' campaign's linked tasks with no changes needed). Skips over empty
' slots to find the nearest real campaign (confirmed with the user) --
' campaigns aren't necessarily packed into consecutive slots, so a
' literal adjacent-slot swap would often silently do nothing.
Public Function HandleCampaignReorder(r As Long, direction As String) As String
    On Error GoTo ErrHandler
    If Not IsMainRow(r) Then
        HandleCampaignReorder = "NOT_MAIN_ROW"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = CampaignsSheet()

    Dim currentSlot As Long
    currentSlot = (r - FIRST_DATA_ROW) \ SLOT_SIZE

    Dim targetSlot As Long
    targetSlot = -1

    Dim s As Long, candRow As Long
    If direction = "up" Then
        For s = currentSlot - 1 To 0 Step -1
            candRow = FIRST_DATA_ROW + s * SLOT_SIZE
            If Len(Trim$(ws.Cells(candRow, COL_CAMPAIGN_NAME).Value & "")) > 0 Then
                targetSlot = s
                Exit For
            End If
        Next s
    Else
        For s = currentSlot + 1 To NUM_SLOTS - 1
            candRow = FIRST_DATA_ROW + s * SLOT_SIZE
            If Len(Trim$(ws.Cells(candRow, COL_CAMPAIGN_NAME).Value & "")) > 0 Then
                targetSlot = s
                Exit For
            End If
        Next s
    End If

    If targetSlot = -1 Then
        HandleCampaignReorder = "NO_TARGET"
        Exit Function
    End If

    Dim targetRow As Long
    targetRow = FIRST_DATA_ROW + targetSlot * SLOT_SIZE

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    SwapCampaignCell ws, r, targetRow, COL_CAMPAIGN_NAME
    SwapCampaignCell ws, r, targetRow, COL_START_DATE
    SwapCampaignCell ws, r, targetRow, COL_END_DATE
    SwapCampaignCell ws, r, targetRow, COL_STATUS
    SwapCampaignCell ws, r, targetRow, COL_NOTES
    SwapCampaignCell ws, r, targetRow, COL_CAMPAIGN_KEY

    RefreshCampaignNameList

    ' User-reported (§19 Phase 5): if either slot's summary block was
    ' expanded at the time of the swap, its unhidden row COUNT stays
    ' whatever it was for the OLD occupant -- collapse state is tied to
    ' the physical slot, not the campaign identity (by design, see this
    ' function's header comment), but the row count needs to match
    ' whichever campaign is now actually there. ws.Calculate first,
    ' forcing Number of Tasks to reflect the just-swapped CampaignKey
    ' before RefreshExpandedVisibility reads it.
    ws.Calculate
    RefreshExpandedVisibility ws, r
    RefreshExpandedVisibility ws, targetRow

    Application.EnableEvents = prevEvents

    HandleCampaignReorder = IIf(direction = "up", "MOVED_UP", "MOVED_DOWN")
    Exit Function
ErrHandler:
    HandleCampaignReorder = "ERROR " & Err.Number & ": " & Err.Description
End Function

Private Sub SwapCampaignCell(ws As Worksheet, row1 As Long, row2 As Long, col As Long)
    Dim tmp As Variant
    tmp = ws.Cells(row1, col).Value
    ws.Cells(row1, col).Value = ws.Cells(row2, col).Value
    ws.Cells(row2, col).Value = tmp
End Sub

' Re-hides then re-unhides exactly the right number of a slot's summary
' rows, based on its CURRENT occupant's live Number of Tasks -- a no-op
' if the slot is already collapsed. Used after HandleCampaignReorder
' swaps two slots' identities, so an expanded block immediately reflects
' whichever campaign now actually occupies it, instead of requiring a
' manual collapse/re-expand to notice the row count is stale.
Private Sub RefreshExpandedVisibility(ws As Worksheet, mainRow As Long)
    If ws.Rows(mainRow + 1).Hidden Then Exit Sub

    Dim numTasks As String
    numTasks = CStr(ws.Cells(mainRow, COL_NUM_TASKS).Value)
    Dim visibleCount As Long
    visibleCount = CLng(Val(numTasks))
    If visibleCount > SUMMARY_ROWS_PER_SLOT Then visibleCount = SUMMARY_ROWS_PER_SLOT

    ws.Rows((mainRow + 1) & ":" & (mainRow + 1 + SUMMARY_ROWS_PER_SLOT)).Hidden = True
    If visibleCount > 0 Then
        ws.Rows((mainRow + 1) & ":" & (mainRow + 1 + visibleCount)).Hidden = False
    End If
End Sub

' Toggles the 8-row hidden summary block beneath a main row (§10.2 point
' 3) -- only meaningful when Number of Tasks > 0, mirroring Tasks' own
' Task-column collapse guard (§6.3).
Public Function HandleCampaignDoubleClick(r As Long) As String
    On Error GoTo ErrHandler
    If Not IsMainRow(r) Then
        HandleCampaignDoubleClick = "NOT_MAIN_ROW"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = CampaignsSheet()

    Dim numTasks As String
    numTasks = CStr(ws.Cells(r, COL_NUM_TASKS).Value)
    If Len(numTasks) = 0 Or numTasks = "0" Then
        HandleCampaignDoubleClick = "NO_LINKED_TASKS"
        Exit Function
    End If

    Dim isCurrentlyHidden As Boolean
    isCurrentlyHidden = ws.Rows(r + 1).Hidden

    If isCurrentlyHidden Then
        ' Expanding: only unhide as many summary rows as are actually
        ' linked (PLAN.md §19 Phase 5, user-requested) -- the "+N more"
        ' overflow row (row 8) already communicates anything past 8, so a
        ' campaign with e.g. 2 linked tasks shows exactly 2 rows, not 8
        ' blank ones. Keeps the fixed 8-row-per-slot architecture (§10.2
        ' point 4) untouched -- this only changes how many of the
        ' existing rows get unhidden, not the row count itself.
        ' Row r+1 is the mini header row (always shown together with
        ' whatever data rows are visible); data rows start at r+2.
        Dim visibleCount As Long
        visibleCount = CLng(Val(numTasks))
        If visibleCount > SUMMARY_ROWS_PER_SLOT Then visibleCount = SUMMARY_ROWS_PER_SLOT
        ws.Rows((r + 1) & ":" & (r + 1 + visibleCount)).Hidden = False
    Else
        ' Collapsing: always re-hide the full header + 8-data-row block,
        ' regardless of how many were actually shown, so the slot returns
        ' to a clean fully-hidden state every time.
        ws.Rows((r + 1) & ":" & (r + 1 + SUMMARY_ROWS_PER_SLOT)).Hidden = True
    End If

    HandleCampaignDoubleClick = IIf(isCurrentlyHidden, "EXPANDED", "COLLAPSED")
    Exit Function
ErrHandler:
    HandleCampaignDoubleClick = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function HandleStatusChange(r As Long) As String
    On Error GoTo ErrHandler
    If Not IsMainRow(r) Then
        HandleStatusChange = "OK"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = CampaignsSheet()

    If Trim$(ws.Cells(r, COL_STATUS).Value & "") = "Completed" Then
        ArchiveCampaign ws, r
        HandleStatusChange = "COMPLETED"
    Else
        HandleStatusChange = "OK"
    End If
    Exit Function
ErrHandler:
    HandleStatusChange = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Reuses "Campaigns Archive"'s pre-allocated first row when it's still
' genuinely untouched, instead of always ListRows.Add-ing a new one --
' same "minimum valid size" convention as Time Log/Archive/Schedule Data
' (PLAN.md §19 Phase 1/5 correction notes), and the exact same bug class
' as Tasks' own Archive sheet (modTasks.bas's NextArchiveRow): a raw
' ListRows.Add always grows the table, leaving that first pre-built row
' permanently blank -- visible as a gap between the header and the first
' real archived campaign.
Private Function NextCampaignArchiveRow(archTbl As ListObject) As ListRow
    If archTbl.ListRows.Count = 1 Then
        Dim firstRow As ListRow
        Set firstRow = archTbl.ListRows(1)
        If Len(Trim$(firstRow.Range(1, 1).Value & "")) = 0 Then
            Set NextCampaignArchiveRow = firstRow
            Exit Function
        End If
    End If
    Set NextCampaignArchiveRow = archTbl.ListRows.Add
End Function

' The archiving cascade (§10.4): copy the main row's static values to
' Campaigns Archive, stamp Completed At, freeze Progress, then clear the
' slot's user-entered cells (including CampaignKey, so a later campaign
' typed into this reused slot re-triggers HandleCampaignNameChange's
' bootstrap) and collapse its summary block. Skips a slot with no
' CampaignKey yet -- can't archive a campaign that was never really
' created (defensive; not expected in normal use since Status can't
' meaningfully be "Completed" before a Campaign Name/key exists).
Private Sub ArchiveCampaign(ws As Worksheet, mainRow As Long)
    If Len(Trim$(ws.Cells(mainRow, COL_CAMPAIGN_KEY).Value & "")) = 0 Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim completedAt As Date
    completedAt = Now()

    Dim archiveWs As Worksheet
    Set archiveWs = ThisWorkbook.Sheets("Campaigns Archive")
    archiveWs.Unprotect

    Dim archTbl As ListObject
    Set archTbl = archiveWs.ListObjects("CampaignsArchiveTable")
    Dim newRow As ListRow
    Set newRow = NextCampaignArchiveRow(archTbl)

    newRow.Range(1, 1).Value = ws.Cells(mainRow, COL_CAMPAIGN_NAME).Value
    newRow.Range(1, 2).Value = ws.Cells(mainRow, COL_START_DATE).Value
    newRow.Range(1, 3).Value = ws.Cells(mainRow, COL_END_DATE).Value
    newRow.Range(1, 4).Value = "Completed"
    newRow.Range(1, 5).Value = ws.Cells(mainRow, COL_PROGRESS).Value   ' frozen static snapshot
    newRow.Range(1, 6).Value = ws.Cells(mainRow, COL_NOTES).Value
    newRow.Range(1, 7).Value = completedAt
    newRow.Range(1, 8).Value = ws.Cells(mainRow, COL_CAMPAIGN_KEY).Value

    archiveWs.Protect UserInterfaceOnly:=True

    ' Clear this slot's user-entered cells for reuse -- must happen BEFORE
    ' CompactCampaignSlots below, so its scan correctly excludes this now-
    ' archived campaign. Number of Tasks/Progress/ReminderDaysLeft/
    ' ReminderSortKey are all position-relative formulas (same self-
    ' healing pattern as Tasks' SortKey/Number of Subtasks, PLAN.md §19
    ' Phase 1 correction note) -- left untouched, they naturally resolve
    ' to blank/0 once Start/End/CampaignKey are cleared.
    ws.Cells(mainRow, COL_CAMPAIGN_NAME).Value = ""
    ws.Cells(mainRow, COL_START_DATE).Value = ""
    ws.Cells(mainRow, COL_END_DATE).Value = ""
    ws.Cells(mainRow, COL_STATUS).Value = ""
    ws.Cells(mainRow, COL_NOTES).Value = ""
    ws.Cells(mainRow, COL_CAMPAIGN_KEY).Value = ""

    ' Squeezes out the gap this archive just left behind (§19 Phase 5,
    ' user-requested) -- also handles collapsing every slot's summary
    ' block and refreshing the dropdown mirror, so no separate calls are
    ' needed here.
    CompactCampaignSlots

    Application.EnableEvents = prevEvents
End Sub
