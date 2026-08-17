# Progress / where we left off

Quick-glance status. For full technical detail on any item, see `PLAN.md`
§20 ("Debugging log, post-handoff") and §21 ("V2 → V3 migration") — this
file just tracks *what's pending* and *what to do next*.

**Last session:** 2026-08-17. **Delivery model changed this session**
(PLAN.md §21.1): the user's old laptop file ("V2") hit Tasks' row
ceiling. Rather than keep hand-patching that live file forever (the old
Bug #1–#10 workflow), they asked for a freshly rebuilt workbook ("V3")
plus a self-service migration tool that copies real data across
**entirely on their own laptop** (confidential/internal-network-only
data, never shared into this session).

## Current status

- **V3 itself: confirmed working** on the real laptop file. Bug #11
  (Campaigns' Up/Down cells keeping button styling on a freed slot after
  a completion/compaction, PLAN.md §20), the formula-bar fix, the
  Notes-icon TaskKey guards, and everything else checked out. **The old
  Bug #1–#10 "hand-edit V2 in place" backlog is now moot** — a clean V3
  rebuild already has all of those fixes baked in, so there's nothing
  left to walk the user through for them. No need to revisit the old
  Bug #2/#3 Notes-icon question either.
- **The row-extender/Dashboard-scaling audit** (user asked: "if rows can
  grow, everything downstream should keep working") came back clean — no
  gaps found, Bug #10 already covers it (PLAN.md §21.4).
- **`vba/modMigration.bas` (the migration tool itself): open bug,
  unreproduced.** `MigrateFromV2` errored out on the user's real V2 file,
  right after picking it — exact error text not captured (user doesn't
  recall it precisely). They already re-entered their data by hand
  instead (**done — don't suggest redoing it**) and, for now, don't want
  to re-run the tool just to capture the error.
  - Investigated hard this session: rebuilt a much more realistic test
    fixture (a genuine macro-enabled old-schema `.xlsm`, not just plain
    data) and drove the real `MigrateFromV2` code path via a new
    test-only entry point (`modMigration.MigrateFromV2Test`). Completed
    cleanly both times — could not reproduce. Whatever's wrong is
    specific to something about the user's actual file (scale, a data
    irregularity, or something not yet considered).
  - **Next session:** ask if they're willing to retry now — the safe way
    is a **throwaway copy** of V3 (never their real, already-populated
    working file), running `MigrateFromV2` against their real V2 there,
    and reporting back the exact `Err.Number`/`Err.Description` (never
    contains actual data). Without that exact text, further blind fixing
    isn't productive — two realistic repro attempts already failed.

## Repo state (this machine, not yet pushed)

Uncommitted changes across `sheets/*.py`, `vba/*.bas` (including the new
`vba/modMigration.bas`), `inject_vba.py`, `CLAUDE.md`, `PLAN.md`, this
file, plus a rebuilt `Performance Marketing Dashboard.xlsm`/`.xlsx`. Per
standing instruction, **don't commit/push until explicitly asked** — as
of this session the user has confirmed V3 itself works, but the migration
tool is still an open bug, and commit timing hasn't been re-discussed
since.
