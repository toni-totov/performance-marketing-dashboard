# Progress / where we left off

Quick-glance status. For full technical detail on any item, see `PLAN.md`
§20 ("Debugging log, post-handoff") and §21 ("V2 → V3 migration") -- this
file just tracks *what's pending* and *what to do next*.

## CURRENT WORK (2026-10-16): V4 personal-tracker refresh -- PLAN.md §22

User approved the full §22 plan (employee variant). Phases A-H, in order:
A bug fixes -> B Recurring Tasks time/per-subtask schedule/placement ->
C Skipped/Rescheduled + auto-skip -> D Archive reshape -> E Daily Log ->
F Today panel -> G visual pass -> H migration V3->V4. Phase status:

| Phase | Status |
|---|---|
| A | done, COM-verified |
| B | done, COM-verified (+ manager smoke test) |
| C | done, COM-verified |
| D | done, COM-verified |
| E | done, COM-verified (Daily Log sheet) |
| F | done, COM-verified (Today panel) |
| G | done, COM-verified + screenshots |
| H | done, COM-verified -- and §21's old migration bug ROOT-CAUSED + fixed |

**All 8 phases built; full regression (A-H + manager smoke) green.** Both
`Performance Marketing Dashboard.xlsm` and `Task Tracker - Manager.xlsm`
rebuilt from current source. **Committed 2026-10-16 as db345b9 on branch
`v4-refresh`** (not merged into master, not pushed). **Not yet done by the user:** actually migrating their
real V3 file into the new build on their laptop (`MigrateFromPrevious`,
Alt+F8, on a throwaway copy first) and using V4 for real.

**Parked by the user (2026-10-16), don't fix unprompted:** manager
build's Dashboard/Calendar/Campaigns read employee Tasks column positions
(PLAN.md §22.8 "Found, not fixed") -- "for now we won't tackle the
manager sheet." The manager .xlsm was NOT rebuilt after this decision.

**Subtask progress (§22.11): built + COM-verified** after the user asked
for it.

Review screenshots/seed script live in the session scratchpad only
(throwaway). `Task_Tracker_Toni_V2/V3.xlsm` in repo root (gitignored) may
be the user's real files -- never open them without asking.

---

**Previous session:** 2026-08-21. **Manager-variant build: COMPLETE (8 of 8
phases), plus FOUR shipped post-testing feedback rounds (a fifth,
sort-by-deadline, was attempted and reverted -- see round #5 below,
NOT in the current file).** The user's manager saw
the V3 workbook, loved it, and asked for her own tailored copy ("Task
Tracker - Manager"). A full implementation plan was designed, reviewed,
and approved (saved at
`C:\Users\STLNOS7AR\.claude\plans\streamed-dazzling-dream.md` -- still the
source of truth for the *design*, though the sort-direction point below
has since been reversed from what it originally said). Both
`Performance Marketing Dashboard.xlsm` (employee) and `Task Tracker -
Manager.xlsm` (manager) now exist and are COM-verified working. **The
user then actually opened and used the manager build** (not just this
session's own COM tests) and reported 3 issues (fixed, "round #1"), asked
for a 4th feature -- per-PIC color coding -- (done, "round #2", which
also includes a real Excel-Table gotcha AND a CF-banding-overrides-
manual-fill gotcha found along the way), reported the formula bar hidden
again -- a recurrence of an old, supposedly-already-fixed V2/employee bug
("round #3"), then reported PIC colors getting masked the moment a task
became a subtask -- the SAME CF-overrides-manual-fill root cause as round
#2's PIC List bug, recurring a second and third time in Tasks' and
Archive's own depth shading/banding, the latter two caught by proactively
auditing for the pattern rather than waiting to be told ("round #4").
**Worth noting as its own pattern, not just four separate bugs:** CF-fill-
overrides-manual-fill is now a recurring root cause across 4 different
call sites in this feature -- worth remembering as a checklist item for
ANY future manually-colored/user-styled cell in this codebase: check
every CF rule that shares its range before assuming a direct fill will
actually display. **Still not done: the user re-confirming any of this
works for them in their own real use** -- everything below is this
session's own
verification. Also still uncommitted (see "Repo state" below) -- nothing
has been pushed.

## Post-build feedback round #1 (after the user actually used the manager file)

Three issues reported after real use; all fixed and COM-verified this
session (`test_feedback_fixes.py` in scratchpad, plus a full Phase 5/6
worked-example regression rerun to confirm the CascadeComplete
restructuring didn't disturb the counter/dedup logic -- it didn't, same
exact values, just correct order now).

1. **PIC List had no way to manually reorder names.** Added Up/Down
   columns (B/C, PIC Name stays A) -- `sheets/pic_list.py`, new
   `modTasksManager.HandlePICListReorder` (plain adjacent-row swap, no
   grouping/hierarchy the way Tasks' own reorder has to account for) and
   `RefreshPICListButtonStyling` (re-scan-and-reconcile: styles Up/Down
   for every occupied row, `ClearButtonLook`s every vacated one -- PIC
   List has no physical row-delete the way Tasks does, so a cleared name
   would otherwise leave stale button borders behind). `vba/PIC List.bas`
   gained a `Worksheet_BeforeDoubleClick` dispatcher; `Worksheet_Change`
   now also calls the new refresh after `RecompactPICList`.
   `ThisWorkbookManager.bas`'s defensive-backfill call list grew to match.
2. **New tasks sorted to the top instead of staying at the bottom.**
   Root cause: blank-PIC ("unassigned") tasks sorted BEFORE every named
   PIC cluster (the original design, confirmed with the user during
   planning) -- since a brand-new task always starts unassigned, it
   visually jumped to the top. **User confirmed the fix: blank-PIC should
   sort AFTER all PIC clusters instead** -- `vba/modHierarchy.bas`'s
   `SortChildrenBySiblingOrder` now uses a new `ComparePICForSort` helper
   (blank sorts last, named clusters compare alphabetically as before) in
   place of a bare `StrComp`. `usePIC=False` path (every other caller)
   untouched; employee output reconfirmed byte-identical after this
   change.
3. **A completed subtask's ancestor appeared BELOW it in Archive**,
   destroying the visual tree shape. Root cause: `CascadeComplete` wrote
   the completing branch's own rows first, then walked ancestors
   afterward -- since Archive is append-only, the parent's context row
   always landed after (below) its child's completed row. **Fix:**
   collect the ancestor chain first (still incrementing each one's
   Completed-Subtasks counter during that same walk), then write ancestor
   context rows in REVERSE order (root down to immediate parent) BEFORE
   the branch's own rows -- so newly-appended rows now read top-down as
   root -> ... -> immediate parent -> completed branch, matching the
   tree's real shape. An ancestor whose context row already existed (from
   an earlier, unrelated completion) still just updates in place via
   `FindOrCreateArchiveRow` -- only newly-created rows needed this fix.
   **User also asked for (and got) visual tree cues, not just correct
   order:** `sheets/archive.py` gained manager-only depth-shading + "⤷ "
   subtask-prefix conditional formatting (same technique as Tasks' own,
   keyed off Archive's existing hidden Level column), and
   `ArchiveCompletedRow`/`ArchiveParentContextRow` now set the Task
   cell's `IndentLevel` too (mirroring what `RenumberHierarchy` already
   does for Tasks). Employee Archive untouched -- it has no Parent-
   Context mechanism to clarify and wasn't part of what was reported.

**A user's own Excel session was found running mid-verification** (a
`tasklist` check turned up an EXCEL.EXE that survived after a test
cleanup) -- confirmed via a non-destructive `GetActiveObject` inspection
(Visible=True, one workbook open, Saved=False) that it was the user's own
in-progress session on their copy of the manager file, NOT a leftover
automation instance. Left completely untouched, per CLAUDE.md's standing
rule against touching any process not tracked at this session's own
launch -- worth remembering if this comes up again.

## Manager-variant plan summary (see the plan file for full detail)

One shared, parametrized codebase (`build_dashboard.py --variant employee|
manager`) produces both workbooks, rather than forking the repo. Manager's
Tasks gains a **PIC** column (dropdown from a maintained roster, auto-
groups siblings by PIC per-parent, blank-PIC tasks stay ungrouped) and a
**Completed Subtasks** running counter; Archive gains a **RowType**
column (`Completed`/`Parent Context`) so a completed subtask's whole
ancestor chain stays visible even while those ancestors are still open in
Tasks; Timer/Time Spent/Monthly Media Plan are dropped entirely for her;
and the `RenumberHierarchy` performance bug (cell-by-cell COM calls) gets
fixed in the shared engine, benefiting both variants.

## Current status -- implementation phases (of 8 total, see plan file §"Sequencing")

- **Phase 1 (`vba/modHierarchy.bas`) -- DONE, verified via COM.** PIC-aware
  two-key sort added (`ColPIC`/`usePIC`, backward-compatible default `0`),
  plus the bulk-array performance rewrite (`LastHierarchyRow` and
  `RenumberHierarchy`'s read/write passes both went from thousands of
  individual `.Cells().Value` COM calls to a handful of bulk `Range.Value`
  transfers). Live-tested: a 7-row mixed-PIC tree sorts exactly as
  designed (blank-PIC stays ungrouped and ordered by SiblingOrder ahead of
  any PIC cluster; PIC clusters sort alphabetically per-parent, case-
  insensitive; same-named PIC under a different parent never bleeds
  across); `ColPIC=0` regression-tested to reproduce the exact original
  pure-SiblingOrder order (zero behavior change for Recurring
  Tasks/employee Tasks/modMigration.bas). New engine: 300 rows in 0.027s.
- **Phase 2 (`sheets/pic_list.py`) -- DONE (Python/structural half).**
  New "PIC List" sheet, manager-only, growable Table + `PICList` dynamic
  named range (`OFFSET(...,MAX(1,COUNTA(...)),1)` -- the `MAX(1,...)`
  floor avoids a zero-height-OFFSET dropdown error while genuinely empty,
  without needing a permanent placeholder row). Live-verified via COM:
  empty list -> one blank item, no error; 1 and 3 names -> exact range,
  no trailing blank; a gap left by clearing a middle name -> the item
  after the gap silently drops out of the dropdown, confirming exactly
  why recompaction VBA is needed. **The VBA recompaction routine itself
  (`PIC List.bas` + a `RecompactPICList` function) is deferred to Phase 4**
  (it naturally belongs in `modTasksManager.bas`, not yet created) --
  don't forget this when Phase 4 starts.
- **Phase 3 (`sheets/tasks.py` variant refactor) -- DONE, verified.**
  `build(wb, variant="employee"|"manager")`. Employee path confirmed
  byte-identical to pre-refactor output (every cell value/formula/number
  format/alignment/column width/hidden-state/data-validation/defined-name
  diffed against `git show HEAD:sheets/tasks.py`'s output -- zero diffs).
  Manager layout verified structurally correct (A-Y column layout exactly
  as planned, PIC/Completed Subtasks columns in the right places with the
  right formulas-or-blanks, PIC dropdown wired to `=PICList`).
- **Phase 4 (`vba/TasksManager.bas` + `vba/modTasksManager.bas`) --
  DONE, verified via COM.** A genuine fork of `modTasks.bas` (per the
  plan, not a runtime-branching shared module), covering: bootstrap of a
  new task row, reparenting, `HandlePICChange` (new -- re-sorts into the
  target PIC cluster via `NextSiblingOrderForPIC`), `HandleSiblingReorder`
  with the PIC-boundary no-op fix, `CascadeDelete`, the row extender's own
  adapted copy (no Timer/Time Spent, PIC dropdown re-applied to new rows,
  Completed Subtasks left blank), and the deferred PIC List recompaction
  routine from Phase 2 (`RecompactPICList`, wired to a new `PIC List.bas`
  sheet code-behind). Live-tested end-to-end on a hand-built fixture
  (build_dashboard.py/inject_vba.py don't support the manager variant
  until Phase 8, so this used a throwaway workbook the same way Phase 1
  tested modHierarchy.bas directly): a 4-child mixed-PIC tree clusters and
  sorts correctly (blank-PIC child first, then alphabetical PIC clusters,
  each internally SiblingOrder-ordered); sibling reorder swaps correctly
  within a PIC cluster and correctly no-ops at a cluster boundary; the row
  extender runs clean end-to-end (allowance 20->520, formulas rewritten
  with the new bound, banner redrawn, PIC dropdown validation added to
  new rows); PIC List recompaction closes a mid-list gap correctly.
  **One real bug caught and fixed during this testing** (via the
  checkpoint-tracing technique, CLAUDE.md): none in the VBA itself this
  round -- the one failure hit (`ExtendTaskRows` erroring 1004) was a test
  *fixture* gap (`CampaignNameList`/`PICList` defined names not present
  in the hand-built test workbook, and Excel's List-validation `.Add`
  throws 1004 against an undefined name), not a real code bug; fixed by
  adding both names to the fixture, not by changing any `.bas` file.
  **Scope note, deliberate:** `Status="Completed"` is NOT wired yet --
  `HandleStatusChange`'s `Case "Completed"` is a documented no-op for now.
  The Completed Subtasks counter + Archive RowType/Parent-Context logic
  (plan §5/§6) needs Archive's manager column layout to exist first, and
  that's Phase 6, not Phase 4 -- building a throwaway `CascadeComplete`
  now that Phase 5 would immediately discard wasn't worth the risk of it
  diverging from Phase 6's actual Archive design. `Status="Delete"` has no
  such dependency and is fully implemented. Also **not yet built:**
  `vba/ThisWorkbookManager.bas` (the `Workbook_Open` defensive call into
  `RecompactPICList`, plan §8) -- not required for Phase 4's own COM
  testing, deferred to whichever later phase actually wires
  `Workbook_Open` for the manager build. **Not verifiable yet either
  way:** the actual PIC-dropdown UX and up/down-click UX in a real,
  fully-built file -- `Worksheet_Change`/`Worksheet_BeforeDoubleClick` are
  `Private Sub`s that can't be driven via `Application.Run` (this
  project's standing limitation), and the manager workbook can't be fully
  built until Phase 8's orchestration lands, so this can only be visually
  confirmed once that exists.
- **Phase 5 (Completed Subtasks counter + Archive RowType/Parent-Context)
  -- DONE, verified via COM.** `modTasksManager.bas` gained
  `CascadeComplete` (Private, restructured from `modTasks.bas`'s version:
  archive loop -> counter/ancestor-walk -> delete bottom-up ->
  RenumberHierarchy -> RefreshButtonStyling), `FindOrCreateArchiveRow`
  (the shared TaskKey-keyed lookup-or-create chokepoint both
  `ArchiveCompletedRow` and `ArchiveParentContextRow` route through --
  this is what makes a Parent-Context row convert to Completed in place
  with no separate conversion logic), and `NextArchiveRow`. Wired into
  `HandleStatusChange`'s `Case "Completed"` (previously a documented
  no-op, see Phase 4's note). Archive's manager column layout
  (ID..RowType, 19 cols) doesn't have a Python builder yet (Phase 6), so
  this was tested the same way Phase 4 tested Tasks' own manager layout:
  a hand-built `ArchiveTable` ListObject matching the column positions
  already fixed by the plan, not a real build.
  **Reconstructed the plan's exact worked example** (task "1" -> "1.2" ->
  {"1.2.1","1.2.2"}, extended with a sibling "1.1" to also cover the
  "two independent branches" case): "1.2.1" completes alone -> Root/Mid12
  both read 1 (matches "1/1"), with fresh Parent-Context Archive rows for
  both; "1.1" (independent branch, same ancestor) completes -> Root's
  *existing* Archive row updates to 2 in place, confirmed exactly one row
  for Root throughout (not a duplicate); "1.2" completes, cascading
  "1.2.2" -> "1.2"'s own archived value comes back 2 (own live 1 +
  descendant "1.2.2" in the same branch) and "1.2.2"'s comes back 0
  (matches "2/0"), "1.2"'s Archive row is the SAME row flipped from
  "Parent Context" to "Completed" in place (via `FindOrCreateArchiveRow`,
  no duplicate), and Root's final counter reads 4 (3 from the plan's pure
  example + 1 for the added "1.1" branch) -- all exact matches.
  **Two real test-script bugs found and fixed along the way (not VBA
  bugs):** (1) the test wasn't disabling `Application.EnableEvents`
  before raw Python cell writes, so the just-injected `TasksManager.bas`
  dispatcher auto-fired every handler a SECOND time on top of the test's
  own explicit `Application.Run` call -- harmless for Phase 4's mostly-
  idempotent handlers, but genuinely destructive once `CascadeComplete`
  (which deletes rows) was in the mix; fixed by disabling events globally
  in the test harness. (2) the test's own `reparent()` helper wrote a
  task's *TaskKey* into "Subtask of" instead of its *visible ID* --
  `HandleSubtaskOfChange` resolves that column via `FindRowByID`, not by
  key, so this silently left every reparented row at the ROOT level
  instead of nested, which in turn made every ancestor-walk test target
  find no ancestors at all. Both were fixed in the throwaway test script
  only; no `.bas` file changed as a result of either.
  **Deliberately not built yet:** `ArchiveManager.bas` (the Archive
  sheet's own code-behind, `Case 14` for Notes vs. employee's `Case 13`)
  -- not needed for Phase 5's COM verification (no double-click handler
  under test), and tightly coupled to Phase 6's actual sheet, so deferred
  there alongside `sheets/archive.py` rather than hand-built ahead of it
  the way Archive's *column layout* was.
- **Phase 6 (`sheets/archive.py` manager layout + `ArchiveManager.bas`) --
  DONE, verified via COM.** `archive.py` gained `build(wb,
  variant="employee"|"manager")` following the same pattern as Phase 3's
  `tasks.py` refactor: employee output confirmed byte-identical via a
  cell/format/table/CF/protection diff against `git show
  HEAD:sheets/archive.py`'s pre-Phase-6 output -- one real diff caught
  this way (a `"Time Spent (hrs)"` column width I'd added that the
  original never set) and fixed. Manager layout (ID, Start Date, Task,
  PIC, Total Subtasks, Completed Subtasks, Subtask of, Project/Category,
  Linked Campaign, Priority, Deadline, Status, Days Left, Notes, Completed
  At, then hidden TaskKey/ParentKey/Level/RowType) confirmed column-for-
  column identical to the positions `modTasksManager.bas`'s
  `ARCHIVE_COL_*` constants already depended on since Phase 5. New
  `ArchiveManager.bas` (sheet code-behind) mirrors `Archive.bas` with
  `Case 14` for Notes instead of employee's `Case 13`.
  **Stronger validation than Phase 5's own hand-built fixture:**
  rebuilt a full REAL manager workbook stack this time (`tasks.py`,
  `pic_list.py`, `archive.py`, `campaigns.py`, `campaigns_archive.py`, all
  via their actual builders, not hand-rolled sheets) and reran Phase 5's
  exact worked example against it end-to-end -- same exact results (1/1
  -> independent-branch update to 2, one row, no duplicate -> 2/0 ->
  final ancestor total 4), this time exercising the real `ArchiveTable`'s
  actual starting shape (one blank template row) rather than an assumed
  one. Also reran `ExtendTaskRows` against the real (1500-row) manager
  Tasks sheet successfully. No VBA changes were needed as a result --
  Phase 5's logic held up unchanged against the real structure it was
  designed for ahead of time.
- **Phase 7 (`sheets/dashboard.py` variant handling) -- DONE, verified via
  COM.** `build(wb, variant=...)` now threads through `_build_kpi_strip`,
  `_build_yearly_monthly_performance`, `_build_performance_charts`, and
  `_build_daily_weekly_productivity` -- everything else (Upcoming Tasks,
  Campaign Reminders, Campaign Metrics Summary) is untouched and keeps its
  exact same hardcoded row position in both variants, per the plan's own
  "accept a small cosmetic gap" call. Employee output confirmed byte-
  identical via the same cell/format/table/CF/chart-anchor diff convention
  as Phases 3/6 (zero diffs). Manager changes, all per plan §7: "Total
  Hours Invested" KPI card dropped, remaining 5 cards recompacted to stay
  contiguous (A/D/G/J/M -- previously A/D/G/J/M/P) -- the Overdue-Tasks
  red-fill and Current-Campaign-Ends warning-border conditional-formatting
  rules, previously hardcoded to columns J/M, now derive their column
  dynamically from the active card list, since those two cards' positions
  shift between variants; "Completed Tasks (Total)" filters on
  `ArchiveTable[RowType]="Completed"`; Yearly Monthly Performance narrows
  to 2 columns (Month, Completed -- Hours/HoursRaw dropped) with only the
  one Completed-count chart remaining, and its monthly COUNTIFS also gains
  the RowType filter; Daily/Weekly Productivity drops both "Time Invested"
  rows (no Timer/Time Log for manager) and adds the same RowType filter to
  its two remaining Completed counts.
  **Verified via `Application.Evaluate` against a real manager workbook
  stack** (`tasks.py`, `pic_list.py`, `archive.py`, `campaigns.py`,
  `campaigns_archive.py`, `dashboard.py`, all real builders) seeded with a
  deliberate mix of `RowType="Completed"` and `RowType="Parent Context"`
  Archive rows: the Completed Tasks (Total) KPI, Completed Today,
  Completed This Week, and the current month's Yearly-Monthly Completed
  count all correctly excluded the Parent-Context rows and matched
  independently-computed expected values (the "this week" expectation was
  computed dynamically from Python's own `datetime`, not hardcoded, so the
  test stays correct regardless of what weekday it's rerun on). Also
  confirmed structurally: table narrows to 2 columns, exactly 1 chart
  (not 2), and the 6th KPI card slot (P1) is genuinely empty.
  **Not yet visually confirmed:** the actual KPI-strip/Daily-Weekly panel
  reflow in a real, opened Excel window -- formula correctness and
  structural shape are COM-verified, but true visual layout (card
  spacing, panel gap below the shrunk Daily/Weekly section) can only be
  eyeballed once a real manager file exists to open (Phase 8).
- **Phase 8 (build orchestration) -- DONE, verified via COM. All 8
  phases now complete -- the manager-variant project is finished.**
  New `build_config.py` (`VARIANTS`, `output_xlsx_path`/`output_xlsm_path`
  -- single source of truth both build scripts import from, instead of
  each hardcoding filenames). `build_dashboard.py`: `build(variant=...)`
  + `--variant` argparse flag; manager skips Time Log/Monthly Media Plan,
  builds PIC List before Tasks. `inject_vba.py`: `STANDARD_MODULES`/
  `SHEET_CODE_BEHIND`/`FORM_SPECS` became variant-aware functions +
  `--variant` flag; `modMigration`/`frmExchangeRate` excluded from
  manager. New `vba/ThisWorkbookManager.bas` (plan §8): identical to
  `ThisWorkbook.bas` minus the Timer-left-running check (whose target,
  `modTasks`, doesn't exist in the manager build), plus a defensive
  `RecompactPICList` call on open.
  **A real gap in the approved plan was found and fixed during this
  phase, not just implementation work:** `modRecurringTasks.bas` --
  documented in the plan as "shared, both variants unchanged" -- actually
  wrote generated task instances directly into Tasks using hardcoded
  EMPLOYEE column positions (`modTasks.NextTaskKey`/`TasksColMap`/
  `RefreshButtonStyling`, plus Timer/Time Spent constants), which doesn't
  exist in the manager build and would have either failed to compile or,
  worse, silently corrupted manager Tasks data (writing into what's
  actually the Status/Days Left columns there) the first time a manager
  recurring-task template fired. Root cause: the plan's "shared" claim
  covered the sheet code-behind and the RT-sheet-only hierarchy engine
  correctly, but missed that `ProcessTemplate`/`GenerateInstance` also
  write into TASKS' column layout, which genuinely differs. **Fix:**
  extracted only the genuinely variant-dependent slice (Project/Linked
  Campaign/Linked Campaign Key/Priority/Deadline/Status, plus Timer/Time
  Spent employee-only, plus `NextTaskKey`/`RenumberHierarchy`/
  `RefreshButtonStyling`) into a new logical module
  `modRecurringTasksWriter` -- both variants inject a component literally
  named that (so `modRecurringTasks.bas`'s own qualified calls resolve
  identically either way), sourced from `modRecurringTasksWriter.bas`
  (employee) or `modRecurringTasksWriterManager.bas` (manager). This is
  the same "one qualified name, variant-selected source file" mechanism
  already used for Tasks/Archive's sheet code-behind (`TasksManager.bas`/
  `ArchiveManager.bas`), now extended to a standard module via
  `inject_vba.py`'s `standard_modules()`. Everything else in
  `modRecurringTasks.bas` (RT-sheet scanning, date math, the hierarchy
  engine for Recurring Tasks' own tree) stayed genuinely untouched, since
  none of it ever touched Tasks' column layout to begin with.
  **Verified end-to-end for both variants:**
  - Employee: `build_dashboard.py --variant employee` +
    `inject_vba.py --variant employee` regenerated the committed
    `Performance Marketing Dashboard.xlsm` -- structural content
    (cell-by-cell, all sheets) confirmed **zero diffs** against the
    previously-committed file. A real close->reopen cycle fired
    `Workbook_Open` cleanly; `RegenerateRecurringTasks` (the modified
    code path) ran via `Application.Run` with no exception, confirming
    the whole project still compiles after the `modRecurringTasksWriter`
    split.
  - Manager: `Task Tracker - Manager.xlsx`/`.xlsm` built and opened for
    the **first time ever**. Confirmed the right module split landed
    (`modTasks`/`modMigration`/`modMonthlyMediaPlan`/`frmExchangeRate`
    all absent; `modTasksManager` present). A real end-to-end Recurring-
    Tasks-generates-into-Tasks smoke test (create a template, force its
    watermark back a week, run `ProcessTemplate`) landed a generated
    instance in manager Tasks with every field in the right column and
    the right value (PIC/Notes/Priority genuinely blank, Linked Campaign
    defaulted to the string `"None"`, Status `"Not Started"`, a real
    Deadline/Days Left, TaskKey assigned) -- this is the exact path the
    plan-gap fix above was for, and it came back completely clean. A
    genuine close->reopen cycle fired `Workbook_Open` (and therefore
    `ThisWorkbookManager.bas`) without error.
  - One false alarm during this verification, worth remembering: a first
    pass appeared to show Linked Campaign landing blank instead of
    `"None"` -- traced with checkpoint tracing (per CLAUDE.md's own
    convention) all the way through `WriteVariantSpecificFields` and
    `RenumberHierarchy`, which both showed the correct value at every
    checkpoint. The actual cause was the test script itself: printing a
    COM cell value with a bare `{val}` f-string renders the STRING
    `"None"` identically to Python's own `None` object, and a "confirmation"
    re-check was reading a stale, unmodified copy of the file (the smoke
    test's own `wb.Close(SaveChanges=False)` discards in-memory changes).
    Using `repr()` in the same live session resolved it immediately.
    Nothing in the `.bas` files needed to change as a result -- pure test
    artifact, but a reminder that `None`-as-a-literal-string is a
    particularly easy footgun for exactly this kind of ambiguous print.

**A real gotcha hit and fixed this session, now also documented in
CLAUDE.md's COM-testing lessons** (read that section before writing any
new COM test script): a `.bas` file's `Attribute VB_Name = "..."` line
must be stripped before `CodeModule.AddFromString` when the component's
`.Name` is already set via the API -- leaving it in silently corrupts the
whole VBA project's compiled state, breaking `Application.Run` for
*every* macro in the project, not just the newly-added one. `inject_vba.py`
already does this correctly (`_strip_attribute_line`); a throwaway test
script that doesn't replicate it will see a generic "macro not available"
error that looks exactly like a Trust Center/security problem but isn't.

## Post-build feedback round #2 (per-PIC color coding)

User asked: "would it be better visually to have each PIC with different
color? Is it possible to apply a way for her to change the color of each
PIC?" Recommended and built: a Color swatch cell per name in PIC List
that she fills using Excel's own native Fill Color picker (not a custom
color-picker UI), propagated by VBA to that PIC's cell wherever the name
appears in Tasks and Archive.

**`sheets/pic_list.py`:** new Color column (D). A swatch left "No Fill"
means "no custom color" downstream.

**`vba/modTasksManager.bas`:**
- `PICColorLookup()` -- scans PIC List's Name/Color columns into a
  Dictionary (`UCase(name) -> Long color`), skipping any swatch that's
  still "No Fill" (checked via `Interior.Pattern`, not just reading
  `Interior.Color` blindly -- an unfilled cell's `.Color` still reads
  back *some* number, which would otherwise misreport "no color chosen"
  as "colored black").
- `ApplyPICColorToCell` -- sets a cell's fill to the looked-up color plus
  a black/white font chosen by a simple perceived-brightness check
  (`TextColorForBackground`, standard luma weights) so text stays
  readable against whatever arbitrary color she picks; explicitly clears
  both when there's no assigned color, so a task reassigned away from a
  colored PIC never keeps a stale color.
- Folded into `RefreshButtonStyling` (Tasks' own re-scan-and-reconcile
  pass, already called at every structural-change site) rather than a
  separate scan -- one pass instead of two.
- `ArchiveCompletedRow`/`ArchiveParentContextRow` gained a `picLookup`
  parameter (built once per `CascadeComplete` call, not per row) so a
  newly-archived row's PIC cell picks up her current color immediately,
  no manual refresh needed for *new* archiving.
- New `HandleRefreshPICColorsClick` -- **required** because Excel/VBA has
  no event that fires on a formatting-only change (`Worksheet_Change`
  never sees a Fill Color edit), so a color changed on an *existing*
  roster entry can't propagate automatically. Re-scans ALL of Tasks
  *and* ALL of Archive (not just newly-touched rows). Wired to a new
  "Refresh PIC Colors" banner on PIC List (mirrors Tasks' own row-
  extender banner pattern) and to `ThisWorkbookManager.bas`'s
  `Workbook_Open` as a same-session-independent backfill.

**A real Excel-Table gotcha found and fixed along the way, not just
feature work:** the Color column was originally added *inside*
`PICListTable`'s own `ref` range (alongside Name/Up/Down). Confirmed live
via COM -- with a plain sequence of cell writes, no VBA even
involved -- that an Excel Table auto-expands to include a newly-typed
row, and when it does, it carries the ROW ABOVE's cell formatting into
every column the table spans, fill color included. This meant a
genuinely brand-new PIC would silently inherit whichever color happened
to sit in the roster row above it the moment the table grew to include
it -- confirmed reproducible, and NOT something she could avoid by using
the UI carefully, since it's inherent to how Tables auto-expand.
Name/Up/Down never showed this symptom only because
`RefreshPICListButtonStyling` already unconditionally overwrites Up/Down
on every edit, masking whatever got silently auto-copied there; Color has
no such per-edit overwrite (it's manager-chosen persistent data, not a
fixed glyph VBA can reset to a known value). **Fix:** moved Color
structurally OUTSIDE `PICListTable`'s `ref` -- it still gets a header,
width, row banding, and borders, just isn't part of the ListObject
itself, so the table's own growth never touches column D at all.
Reconfirmed via the same COM repro that this removes the leak entirely.
Worth remembering for any future column added next to a growing Table
elsewhere in this project: cells *outside* a table's `ref` are just
plain cells with no such auto-copy behavior; cells *inside* one inherit
the row-above's format the moment the table expands to include them,
regardless of whether that expansion came from a real keystroke or a
COM-driven `.Value` write.

Full test suite (`test_pic_colors.py` in scratchpad) reran clean after
the fix: Bob's dark-navy swatch -> Tasks PIC cell colored + white text;
Alice's pale-yellow swatch -> Tasks PIC cell colored + black text;
Carol's "No Fill" swatch and a genuinely-unassigned task -> both
correctly uncolored; completing Bob's task carried the color into
Archive automatically; changing Bob's color afterward + clicking
"Refresh PIC Colors" updated BOTH his live Tasks row and his
already-archived row, while Alice's stayed untouched; clearing a task's
PIC assignment correctly cleared its cell's color. The 3 fixes from
round #1 were also regression-rerun clean against this same rebuild, and
employee output reconfirmed byte-identical (pic_list.py is manager-only,
so no employee exposure was expected, but reconfirmed rather than
assumed).

**A SECOND real bug found the same way -- the user actually using it,
same session:** "row 1 I can't color, row 2 colors, row 3 does not, row 4
colors" -- an alternating pattern. Root cause: PIC List's row-banding
conditional-format fill (`cf.row_banding_rule`, alternates every other
row by design via `MOD(ROW()-2,2)=0`) was still scoped across the Color
column too. A conditional-format fill visually overrides a cell's own
manually-set direct fill wherever the CF condition evaluates true -- so
her swatch color was invisible on every "banded" row and fine on the
rows in between, exactly the pattern reported. Distinct root cause from
the earlier table-auto-expand bug (this one is a pure *display* issue,
not a data-corruption one -- her actual `Interior.Color` was presumably
intact underneath the whole time, just visually masked), but the same
underlying lesson: a swatch cell's only job is to show her chosen color
unconditionally, so it can't safely carry ANY conditional formatting of
its own, banding included. **Fix:** row banding now stops at column C;
Color keeps its border/width but no fill-based CF at all. Verified via
the reliable method for this (per CLAUDE.md: check the CF rule's actual
range/definition via `FormatConditions`, not a headless-rendered
display) -- column A still carries the banding rule, column D now
carries zero conditional-format rules. Full `test_pic_colors.py` suite
reran clean again after this fix too.

## Post-build feedback round #3 (formula bar hidden -- a recurrence of an
old, already-"fixed" V2/employee bug, now on the manager build)

User: "the formula bar is hidden... this breaks a little bit everything
and it is annoying when you have another excel opened" -- explicitly
tied it to feedback they'd already given once before, on the V2 file.

Context (already documented at the bottom of both `vba/ThisWorkbook.bas`
and `vba/ThisWorkbookManager.bas`, pre-dating this session): an earlier
modernization pass used to toggle `Application.DisplayFormulaBar` off/on
via `Workbook_Activate`/`Workbook_Deactivate`, and that got reverted
(confirmed present in both files, no such handlers exist) because
toggling it on every window-switch was cancelling the OS clipboard's
copy/"marching ants" state. **What the revert missed:** it made the file
passive instead of corrective. `DisplayFormulaBar` is an Excel-wide
setting, not a per-file one -- it persists across sessions on a machine.
So if ANY other open workbook (most plausibly an older file that predates
the revert and still carries the original toggle code) turns it off,
neither `ThisWorkbook.bas` nor `ThisWorkbookManager.bas` had anything
that would turn it back on, leaving the manager file stuck with a hidden
formula bar despite never itself doing anything wrong.

**Fix:** both files' `Workbook_Open` now does a ONE-TIME
`Application.DisplayFormulaBar = True` assertion. Deliberately not a
re-introduction of the old bug -- the clipboard problem was specifically
about toggling on EVERY activate/deactivate cycle (repeated, as the user
switches windows back and forth); a single assertion that only fires
once, when the file is first opened, structurally cannot reproduce that,
since it never runs again on later window switches.

**Verified via COM for both variants** (`test_formula_bar.py` in
scratchpad): forced `Application.DisplayFormulaBar = False` before even
opening each file (simulating exactly the reported scenario -- another
workbook left it off), then did a genuine open (firing `Workbook_Open`
for real, same convention as every other `Workbook_Open`-only-testable
behavior in this project) and confirmed it read back `True` immediately
after, for both `Task Tracker - Manager.xlsm` and `Performance Marketing
Dashboard.xlsm`. Full regression suite (PIC colors, the 3 round-#1
fixes, employee structural diff) reran clean against this same rebuild.

## Post-build feedback round #4 (PIC color masked on subtasks)

User: "when the tasks are tasks, the PICs color as expected, however,
when they become subtasks, the color gets overridden of the subtask
formatting." Third occurrence of the exact same root-cause class as
round #2's PIC List banding fix and round #3's near-miss: a conditional-
format fill overrides a cell's own manually-set direct fill wherever the
CF's condition is true. This time the offending CF was Tasks' own depth
shading (`_apply_depth_shading` in `sheets/tasks.py`, §6.4 -- "the
subtask formatting" she means) -- applied uniformly across the WHOLE
visible row A:P for every Level >= 1, PIC column included. A top-level
task (Level 0) never triggers the shading rule at all, so its PIC color
showed fine; the instant a task became a subtask (Level >= 1), the
shading rule's condition went true and painted over the PIC cell's own
color.

**Proactively audited for the same pattern elsewhere before fixing just
the one reported instance** (this project's own established convention
-- CLAUDE.md: "check every instance of this pattern project-wide the
first time you fix one") and found TWO more live instances, not yet
reported but certain to surface the same way: Archive's own depth
shading (`_apply_archive_depth_shading`, added round #2) has the
identical bug, and Archive's row banding (`cf.row_banding_rule`, applied
across the whole visible row including PIC) has the SAME bug round #2
already fixed once for PIC List -- just never applied to Archive's own
banding at the time, since PIC coloring in Archive didn't exist yet when
Archive's banding was originally built (Phase 6, before PIC colors were
a feature at all).

**Fix, all three sites:** `sheets/tasks.py`'s `_apply_depth_shading` and
`sheets/archive.py`'s `_apply_archive_depth_shading`/row-banding call
each gained an `exclude_col_letter` parameter (PIC, manager only) that
splits the CF range into two sub-ranges around that one column instead of
applying uniformly -- a new shared `_visible_ranges_excluding` helper in
`archive.py` handles the splitting for both of its call sites.
`exclude_col_letter=None` (employee, and any future non-PIC caller)
reproduces the original single-range behavior exactly, so employee output
stays byte-identical.

**Verified via COM:** a genuine Level-1 subtask's PIC cell now has ZERO
`FormatConditions` at all (same as a top-level task's), confirmed for
both Tasks and Archive, while the Task column's own CF rules are
untouched (6/7 rules respectively). Functionally: assigned a PIC to a
real subtask, confirmed its color displays correctly (not masked) both
live in Tasks and after archiving. Full regression suite (PIC colors,
round-#1 fixes, formula bar, employee structural diff) reran clean.

## Post-build feedback round #5 (sort-by-deadline -- attempted, reverted, NOT shipped)

User asked: "can we sort by deadline? And if true, can we introduce a
button to sort by deadline, so the most important projects become on
top?" Scoped to the manager variant only: a "Deadline ↓" clickable column
header that would re-sort tasks by ascending Deadline (soonest first,
blank last) *within* each existing (ParentKey, PIC) group, without
disturbing PIC clustering.

**Built, then hit an unresolved Excel-crashing bug and was fully reverted
before shipping -- the current `Task Tracker - Manager.xlsm` contains
NONE of this work.** `vba/modTasksManager.bas` gained
`HandleSortByDeadlineClick` (grouped rows by `ParentKey|PIC`, insertion-
sorted each group by Deadline with a `9999999` blank-sort-last sentinel,
matching Dashboard's existing sentinel convention, then rewrote
`SiblingOrder` and called `RenumberHierarchy`); `vba/TasksManager.bas`
gained a `Worksheet_BeforeDoubleClick` case for the header cell;
`sheets/tasks.py`'s manager Deadline header gained a "↓" glyph to mark it
clickable.

**Every version tried crashed the whole Excel process when invoked via
`Application.Run`** -- not a blocking modal dialog (this project's usual
`Application.Run` hang cause, per CLAUDE.md), but an actual process death:
the COM call itself fails immediately with `pywintypes.com_error:
(-2147023170, 'The remote procedure call failed.')`, and a subsequent
`Quit()` fails too with "The RPC server is unavailable" -- both symptoms
of the server process being gone, not stuck. Three rewrites (a
Dictionary-of-Collections grouping, then plain arrays with a separate
sort Sub, then everything fully inlined into one function) all crashed
identically, including a completely unrelated canary function
(`RefreshPICListButtonStyling`) called from the same freshly-injected
project -- meaning the corruption/crash wasn't isolated to the new
function's own logic. Root cause was **not found** before the session
ended. Leading unconfirmed suspicion: something in the interaction
between this new `SiblingOrder`-rewriting logic and `modHierarchy.bas`'s
recursive tree walk in `RenumberHierarchy` (e.g. a duplicate/colliding
`SiblingOrder` across groups, or a corrupted `ParentKey` chain) driving
runaway native recursion -- a stack overflow would explain a hard process
crash in a way no VBA-level `On Error` handler could ever catch, unlike
every other documented `Application.Run` failure mode in this project.
This was never confirmed, only reasoned toward, when the feature was
pulled instead of continuing to debug it.

**Decision: not needed right now.** The user wants to send this build to
their manager today for a few days' trial feedback, so rather than keep
debugging a crash on a nice-to-have feature, all three files were
reverted to their pre-feature state (`sheets/tasks.py`'s header text back
to plain "Deadline", the dispatcher case removed from
`vba/TasksManager.bas`, `HandleSortByDeadlineClick` and its `FileTrace`
debug scaffolding deleted entirely from `vba/modTasksManager.bas`), and
`Task Tracker - Manager.xlsm` was rebuilt + re-injected fresh from that
reverted source (confirmed via grep: zero remaining references to
`HandleSortByDeadlineClick`/`sort_trace` anywhere under `vba/`). All four
earlier feedback-round fixes (PIC coloring, alternating-row banding,
subtask masking, formula bar) are untouched and remain in this rebuild,
since they live in files this revert didn't touch.

**If sort-by-deadline comes back as a request:** don't restart from the
inlined version above -- the crash was never actually diagnosed, so
reusing that code would just reintroduce the same unresolved bug.
Start by writing a minimal isolated repro (a hand-built fixture with 2-3
rows, no PIC assignment) that calls `RenumberHierarchy` directly with a
deliberately-duplicated `SiblingOrder` within one group, to test the
"runaway recursion in the tree walk" theory in isolation before touching
the sort logic itself again.

## Old open item, still unresolved (separate from the manager-variant work)

- **RESOLVED 2026-10-16 (PLAN.md §22.9):** root cause reproduced and
  fixed -- protected Archive/Campaigns Archive refused `ListRows.Add` on
  the 2nd archived row. The note below is kept for history only.
- **`vba/modMigration.bas` (V2->V3 migration tool): open bug,
  unreproduced.** `MigrateFromV2` errored on the user's real V2 file once;
  they re-entered the data by hand instead (done, don't suggest redoing
  it) and don't want to re-run the tool just to capture the error yet.
  Two realistic repro attempts (including a genuine macro-enabled test
  fixture) came back clean. If they're ever willing to retry, the safe
  way is a throwaway copy of V3 (never the real working file), and
  reporting back the exact `Err.Number`/`Err.Description`.

## Repo state (this machine, not yet pushed)

**All 8 phases complete, plus all four SHIPPED post-build feedback rounds
above -- working tree is now a coherent, working state (both variants
build and run clean), but nothing has been committed or pushed yet.** Per
standing instruction, don't commit/push until explicitly asked. The
current `Task Tracker - Manager.xlsm` was freshly rebuilt + re-injected
2026-08-21 after round #5's revert -- it reflects rounds #1-#4 only, not
the reverted sort-by-deadline attempt.

Modified (tracked): `CLAUDE.md`, `PROGRESS.md`, `.gitignore` (added the
manager `.xlsx` intermediate), `Performance Marketing Dashboard.xlsm`
(regenerated, confirmed zero structural diff from the previously-committed
version -- reconfirmed again after every feedback round too),
`build_dashboard.py`, `inject_vba.py`, `sheets/archive.py` (Phase 6 +
round-#1 depth-shading/prefix/indentation + round-#4 PIC-exclusion on
both depth shading and row banding), `sheets/dashboard.py`,
`sheets/tasks.py` (round-#4 `_apply_depth_shading`'s new
`exclude_col_letter` param), `vba/ThisWorkbook.bas` (round-#3
`DisplayFormulaBar` fix), `vba/modHierarchy.bas` (Phase 1 + round-#1
`ComparePICForSort`), `vba/modRecurringTasks.bas` (the Phase 8 plan-gap
fix).

New/untracked: `build_config.py`, `sheets/pic_list.py` (round-#1 Up/Down
columns + round-#2 Color column, kept structurally outside
`PICListTable` AND excluded from row banding -- see round #2's own two
correction notes), `vba/modTasksManager.bas` (round-#1
`HandlePICListReorder`/`RefreshPICListButtonStyling`/`CascadeComplete`
reorder + round-#2 `PICColorLookup`/`ApplyPICColorToCell`/
`TextColorForBackground`/`HandleRefreshPICColorsClick`),
`vba/TasksManager.bas`, `vba/PIC List.bas` (round-#1
`Worksheet_BeforeDoubleClick` + round-#2 refresh-banner case),
`vba/ArchiveManager.bas`, `vba/ThisWorkbookManager.bas` (round-#1 +
round-#2 defensive-refresh calls + round-#3 `DisplayFormulaBar` fix,
mirroring `ThisWorkbook.bas`), `vba/modRecurringTasksWriter.bas`,
`vba/modRecurringTasksWriterManager.bas`, `Task Tracker - Manager.xlsm`
(the new manager deliverable, mirroring how `Performance Marketing
Dashboard.xlsm` is committed for employee).

**Not mine, do not touch:** `Task Tracker - Manu.xlsm` (untracked, repo
root) -- the user's own saved copy (truncated filename, presumably from
their own Save As), containing whatever they've been testing with. Left
completely alone all session, same as their live Excel process was
earlier.

Also gitignored (throwaway intermediates, same convention as the
employee `.xlsx`): `Task Tracker - Manager.xlsx` -- added to `.gitignore`
this session.

**Next step is a commit decision, not more implementation** -- when
asked, stage everything above (both new `.bas`/`.py` files and the two
regenerated workbook artifacts) as one commit, since the manager-variant
feature only makes sense as a whole.
