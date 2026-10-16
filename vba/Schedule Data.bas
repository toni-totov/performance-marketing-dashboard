Attribute VB_Name = "Schedule Data"
' NOT a standard module -- this file's content is injected directly into
' the Schedule Data WORKSHEET's own code module by Phase 4's
' inject_vba.py (see PLAN.md §18's sheet-named-.bas convention). All
' editing normally happens via frmScheduleDay (§13.3), but this sheet is
' deliberately left visible and unprotected "in case manual inspection or
' a bulk fix is ever needed" (schedule_data.py's own docstring) -- this
' single Worksheet_Change handler is what makes a manual edit to the
' Meetings table actually reschedule reminders (§13.5), rather than
' silently going stale until the next Workbook_Open. The popup's own
' writes run with EnableEvents off, so they call modSchedule.ScheduleReminders
' directly instead of relying on this handler, which never fires for them.

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column >= 4 And Target.Column <= 10 Then   ' the Meetings table, D:J
        modSchedule.ScheduleReminders
    End If
    ' Schedule's day labels read both tables (V4, PLAN.md §22.8).
    modSchedule.RenderScheduleBadges
End Sub
