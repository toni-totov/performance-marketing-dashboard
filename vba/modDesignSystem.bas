Attribute VB_Name = "modDesignSystem"
Option Explicit

' Shared color/typography/icon constants -- the VBA mirror of
' sheets/design_system.py, itself sourced from PLAN.md §17 (Visual Design
' System). VBA-driven formatting (dynamic fills, MsgBox/popup text) reads
' from here so it stays consistent with openpyxl-applied formatting.
'
' Color constants are RGB(r,g,b) calls, not raw &H hex literals -- VBA's
' Long color values are stored in BGR byte order, so a literal hex string
' copy-pasted from the palette would silently produce the wrong color.
' RGB() handles the byte-order conversion correctly. These values are
' generated from sheets/design_system.py, not hand-transcribed, to avoid
' the two files drifting out of sync.

' --- Brand / chrome accent (§17.1, §17.17) --------------------------------
Public Const COLOR_BRAND_ACCENT As Long = &H794E1F        ' #1F4E79 corporate navy (was the earlier brand-blue candidate #1428A0)
Public Const COLOR_BRAND_ACCENT_TEXT As Long = &HFFFFFF   ' #FFFFFF

' --- Categorical data palette (§17.1, §17.2), fixed order, never cycled --
Public Const COLOR_CATEGORICAL_1 As Long = &HD6782A       ' #2A78D6 blue
Public Const COLOR_CATEGORICAL_2 As Long = &H7AAF1B       ' #1BAF7A aqua
Public Const COLOR_CATEGORICAL_3 As Long = &H00A1ED       ' #EDA100 yellow
Public Const COLOR_CATEGORICAL_4 As Long = &H008300       ' #008300 green
Public Const COLOR_CATEGORICAL_5 As Long = &HA73A4A       ' #4A3AA7 violet
Public Const COLOR_CATEGORICAL_6 As Long = &H4849E3       ' #E34948 red
Public Const COLOR_CATEGORICAL_7 As Long = &HA47BE8       ' #E87BA4 magenta
Public Const COLOR_CATEGORICAL_8 As Long = &H3468EB       ' #EB6834 orange

' --- Status / severity palette (§17.3), fixed, reserved meaning ----------
Public Const COLOR_STATUS_GOOD As Long = &H0CA30C         ' #0CA30C
Public Const COLOR_STATUS_WARNING As Long = &H19B2FA      ' #FAB219
Public Const COLOR_STATUS_SERIOUS As Long = &H5A83EC      ' #EC835A
Public Const COLOR_STATUS_CRITICAL As Long = &H3B3BD0     ' #D03B3B

Public Const COLOR_PRIORITY_LOW As Long = COLOR_STATUS_GOOD
Public Const COLOR_PRIORITY_MEDIUM As Long = COLOR_STATUS_WARNING
Public Const COLOR_PRIORITY_HIGH As Long = COLOR_STATUS_CRITICAL

Public Const COLOR_URGENCY_OVERDUE As Long = COLOR_STATUS_CRITICAL
Public Const COLOR_URGENCY_DUE_SOON As Long = COLOR_STATUS_WARNING
Public Const COLOR_URGENCY_ON_TRACK As Long = COLOR_STATUS_GOOD

' --- Chrome & ink (§17.6) --------------------------------------------------
Public Const COLOR_INK_PRIMARY As Long = &HB0B0B          ' #0B0B0B
Public Const COLOR_INK_SECONDARY As Long = &H4E5152       ' #52514E
Public Const COLOR_INK_MUTED As Long = &H818789           ' #898781
Public Const COLOR_GRIDLINE As Long = &HD9E0E1            ' #E1E0D9
Public Const COLOR_PAGE_BG As Long = &HF7F9F9             ' #F9F9F7
Public Const COLOR_CARD_SURFACE As Long = &HFBFCFC        ' #FCFCFB

' --- Workflow-status mini-scheme (§17.4) -----------------------------------
Public Const COLOR_WORKFLOW_NOT_STARTED As Long = COLOR_INK_MUTED
Public Const COLOR_WORKFLOW_IN_MOTION As Long = COLOR_CATEGORICAL_1
Public Const COLOR_WORKFLOW_DONE As Long = COLOR_STATUS_GOOD

Public Const COLOR_WORKFLOW_ACTIVE As Long = COLOR_STATUS_GOOD
Public Const COLOR_WORKFLOW_INACTIVE As Long = COLOR_INK_MUTED

Public Const COLOR_ACCESS_YES As Long = COLOR_STATUS_GOOD
Public Const COLOR_ACCESS_NO As Long = COLOR_INK_MUTED

' --- Diverging pair (§17.5) -------------------------------------------------
Public Const COLOR_DIVERGING_POSITIVE As Long = COLOR_CATEGORICAL_1   ' blue -- early
Public Const COLOR_DIVERGING_NEGATIVE As Long = COLOR_CATEGORICAL_6   ' red -- late
Public Const COLOR_DIVERGING_MIDPOINT As Long = &HECEFF0              ' #F0EFEC

' --- Typography (§17.7) -----------------------------------------------------
Public Const FONT_NAME As String = "Segoe UI"
Public Const FONT_SIZE_TITLE As Integer = 15
Public Const FONT_SIZE_KPI As Integer = 22
Public Const FONT_SIZE_HEADER As Integer = 11     ' VBA Font.Size is whole points; §17.7's 10.5pt rounds up
Public Const FONT_SIZE_BODY As Integer = 11
Public Const FONT_SIZE_CAPTION As Integer = 9

' --- Iconography (§17.8) ----------------------------------------------------
' Built via ChrW() from each glyph's numeric Unicode codepoint(s), NOT as
' raw string literals -- confirmed via live testing (PLAN.md §19 Phase 5)
' that a Unicode character embedded directly as a VBA source-code string
' literal gets silently corrupted to "?" the moment the module compiles,
' on a system whose ANSI codepage can't represent it (this Bulgarian-
' locale/Windows-1251 build machine can't represent any of these dingbat/
' emoji glyphs). This is a VBA/VBE source-text limitation, not a runtime
' one -- every OTHER icon in this workbook is safe because openpyxl writes
' it directly into cell XML (genuine Unicode) at build time; these are the
' only glyphs that ever flow through VBA source text via AddFromString.
' ChrW() takes a numeric codepoint (itself plain ASCII in the VBA source,
' so immune to the codepage issue) and constructs the real Unicode
' character at RUNTIME, bypassing the compile-time corruption entirely.
' A supplementary-plane character (anything above the Basic Multilingual
' Plane, i.e. most emoji) needs its UTF-16 surrogate pair -- two ChrW()
' calls concatenated, not one.
Public Function ICON_NOT_STARTED() As String
    ICON_NOT_STARTED = ChrW(11036)                        ' U+2B1C white large square
End Function
Public Function ICON_IN_PROGRESS() As String
    ICON_IN_PROGRESS = ChrW(55357) & ChrW(56629)           ' U+1F535 blue circle
End Function
Public Function ICON_COMPLETED() As String
    ICON_COMPLETED = ChrW(9989)                            ' U+2705 check mark
End Function
Public Function ICON_TIMER_RUNNING() As String
    ICON_TIMER_RUNNING = ChrW(9654)                        ' U+25B6 play triangle
End Function
Public Function ICON_NOTES() As String
    ICON_NOTES = ChrW(55357) & ChrW(56541)                 ' U+1F4DD memo
End Function
Public Function ICON_UP() As String
    ICON_UP = ChrW(8593)                                   ' U+2191
End Function
Public Function ICON_DOWN() As String
    ICON_DOWN = ChrW(8595)                                 ' U+2193
End Function
Public Function ICON_NAV_PREV() As String
    ICON_NAV_PREV = ChrW(9664)                              ' U+25C0
End Function
Public Function ICON_NAV_NEXT() As String
    ICON_NAV_NEXT = ChrW(9654)                              ' U+25B6, same glyph as ICON_TIMER_RUNNING;
                                                              ' never on the same sheet (see design_system.py)
End Function
Public Function ICON_REFRESH() As String
    ICON_REFRESH = ChrW(55357) & ChrW(56580)                ' U+1F504
End Function
Public Function ICON_CONFIRM() As String
    ICON_CONFIRM = ChrW(10003)                              ' U+2713
End Function
Public Function ICON_CANCEL() As String
    ICON_CANCEL = ChrW(10005)                               ' U+2715
End Function
Public Function ICON_EXCHANGE_RATE() As String
    ICON_EXCHANGE_RATE = ChrW(55357) & ChrW(56497)          ' U+1F4B1
End Function

' --- Locale-proof month/weekday names (PLAN.md §19 Phase 5, found live) ----
' VBA's Format()/date functions render "mmmm"/"dddd" tokens in the
' SYSTEM locale, same underlying issue as Excel's TEXT() on the
' worksheet side (confirmed live: "July" renders as "юли" on this
' Bulgarian-locale machine) -- a hardcoded English lookup sidesteps it
' entirely, same reasoning as sheets/week_grid.py's INDEX()-based
' formulas on the worksheet side.
Public Function EnglishMonthName(monthNum As Long) As String
    Dim names() As String
    names = Split("January,February,March,April,May,June,July,August,September,October,November,December", ",")
    EnglishMonthName = names(monthNum - 1)
End Function

Public Function EnglishWeekdayName(d As Date) As String
    Dim names() As String
    names = Split("Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday", ",")
    EnglishWeekdayName = names(Weekday(d, vbMonday) - 1)
End Function

' Mirrors sheets/design_system.py's apply_button_look() -- but called
' dynamically per-row from VBA (Tasks'/Recurring Tasks' HandleTaskNameChange,
' Campaigns' HandleCampaignNameChange) rather than baked in across a whole
' sheet's pre-allocated row allowance at build time. §19 Phase 6,
' user-reported: a build-time version styled every one of Tasks' 500
' pre-formatted rows identically, so the button borders visibly continued
' downward through hundreds of rows that don't hold a task yet -- "these
' should appear, when I add new task," not before. Called once per button
' cell at the exact moment a row becomes a real task/campaign (the same
' isNewRow-guarded block that already sets that cell's icon text).
' Hides Excel's green "number stored as text" triangle on an ID-style cell
' (V4, PLAN.md §22.8) -- IDs are text on purpose ("1.10" must not become
' 1.1). The ignore flag is per cell AND is cleared whenever the cell's
' value is rewritten (confirmed via COM), so a build-time setting can't
' survive RenumberHierarchy -- callers re-apply it after every renumber,
' alongside the button restyling they already do per row. Single cell
' only: Range.Errors fails on a multi-cell range.
Public Sub HideNumberAsTextFlag(targetCell As Range)
    On Error Resume Next
    targetCell.Errors(xlNumberAsText).Ignore = True
End Sub

' Lets the user TYPE dates as 17.10.2026 (optionally "17.10.2026 14:30",
' "17.10.26" or just "17.10" for this year) on any machine (user-requested
' 2026-10-17). Every date cell already DISPLAYS as DD.MM.YYYY, but how
' Excel parses typed input follows the Windows region settings, not the
' cell's format -- on a US-locale laptop "17.10.2026" stays plain text and
' every formula reading the cell breaks. Called first thing from each
' sheet's Worksheet_Change: any date-formatted cell (NumberFormat starting
' "DD.MM.YYYY", so no per-sheet column lists to keep in sync) that ended
' up holding text in day.month.year order is rewritten as a real date.
' Text that isn't a valid date is left alone. Multi-cell safe (paste/fill).
Public Sub NormalizeDateInput(ByVal Target As Range)
    On Error GoTo Done
    If Target.CountLarge > 5000 Then Exit Sub
    Dim c As Range, d As Variant
    For Each c In Target.Cells
        If VarType(c.Value) = vbString Then
            If UCase$(Left$(Replace(c.NumberFormat, "\", ""), 10)) = "DD.MM.YYYY" Then   ' may read back escaped as dd\.mm\.yyyy
                d = ParseDayFirstDate(CStr(c.Value))
                If Not IsEmpty(d) Then
                    Application.EnableEvents = False
                    c.Value = d
                    Application.EnableEvents = True
                End If
            End If
        End If
    Next c
Done:
    Application.EnableEvents = True
End Sub

' "d.m.yyyy[ h:mm]" (also "/" or "-" as the separator, 2-digit year, or
' no year) -> a Date, or Empty when the text isn't a real calendar date.
' Built from numbers via DateSerial/TimeSerial, never CDate/DateValue,
' which would re-apply the very locale rules this exists to bypass.
Public Function ParseDayFirstDate(ByVal s As String) As Variant
    On Error GoTo Invalid   ' e.g. an overflowing digit run -- just "not a date"
    Dim datePart As String, timePart As String, p As Long
    s = Trim$(s)
    p = InStr(s, " ")
    If p > 0 Then
        datePart = Left$(s, p - 1)
        timePart = Trim$(Mid$(s, p + 1))
    Else
        datePart = s
    End If
    datePart = Replace(Replace(datePart, "/", "."), "-", ".")
    If Right$(datePart, 1) = "." Then datePart = Left$(datePart, Len(datePart) - 1)

    Dim bits() As String
    bits = Split(datePart, ".")
    If UBound(bits) < 1 Or UBound(bits) > 2 Then Exit Function
    Dim i As Long
    For i = 0 To UBound(bits)
        If Len(bits(i)) = 0 Or Not bits(i) Like String$(Len(bits(i)), "#") Then Exit Function
    Next i

    Dim dd As Long, mm As Long, yy As Long
    dd = CLng(bits(0)): mm = CLng(bits(1))
    If UBound(bits) = 2 Then
        yy = CLng(bits(2))
        If Len(bits(2)) <= 2 Then
            yy = yy + 2000
        ElseIf Len(bits(2)) <> 4 Then
            Exit Function
        End If
    Else
        yy = Year(Date)
    End If
    If mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then Exit Function
    Dim result As Date
    result = DateSerial(yy, mm, dd)
    If Day(result) <> dd Then Exit Function   ' e.g. 31.02 rolled into March

    If Len(timePart) > 0 Then
        Dim hm() As String
        hm = Split(timePart, ":")
        If UBound(hm) <> 1 Then Exit Function
        If Len(hm(0)) = 0 Or Len(hm(1)) = 0 Then Exit Function
        If Not hm(0) Like String$(Len(hm(0)), "#") Or Not hm(1) Like String$(Len(hm(1)), "#") Then Exit Function
        If CLng(hm(0)) > 23 Or CLng(hm(1)) > 59 Then Exit Function
        result = result + TimeSerial(CLng(hm(0)), CLng(hm(1)), 0)
    End If
    ParseDayFirstDate = result
    Exit Function
Invalid:
    ParseDayFirstDate = Empty
End Function

Public Sub ApplyButtonLook(targetCell As Range)
    With targetCell
        .Interior.Color = COLOR_CARD_SURFACE
        Dim edge As Variant
        For Each edge In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            With .Borders(edge)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = COLOR_INK_MUTED
            End With
        Next edge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

' The undo of ApplyButtonLook -- restores a cell to the plain, un-occupied
' "empty slot" look a build-time cell has before any row/slot ever occupies
' it: no fill, and only the bottom hairline border every table-body cell
' gets from sheets/design_system.py's apply_table_hairline_borders (style
' "hair", COLOR_GRIDLINE, bottom edge only -- top/left/right cleared).
' Needed anywhere a row/slot can go from occupied back to empty WITHOUT a
' real Rows.Delete carrying the surrounding blank formatting up with it
' (Campaigns' fixed-slot CompactCampaignSlots, PLAN.md §20 user-reported --
' unlike Tasks'/Recurring Tasks' hierarchy engine, which physically deletes
' rows via CascadeDelete, so a vacated row is never left with stale button
' styling in the first place).
Public Sub ClearButtonLook(targetCell As Range)
    With targetCell
        .Interior.Pattern = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .Color = COLOR_GRIDLINE
        End With
    End With
End Sub
