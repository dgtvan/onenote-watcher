# Requirements Review — Does the Solution Meet the Brief?

Every requirement you stated, and an honest verdict on whether the app meets it. Where something is
only partly met, the gap is named rather than smoothed over. Design rationale is in
[detection-design.md](detection-design.md); the fail-closed contract in [fail-closed.md](fail-closed.md).

## Traceability matrix

| # | Requirement | Verdict | How it is met | Gap / note |
|---|---|---|---|---|
| R1 | Background app, tray icon | ✅ Full | WinForms `NotifyIcon`, no main window, single-instance, run-at-login (HKCU Run) | — |
| R2 | Icon turns **red + pulsing** on any error | ✅ Full | Tray state machine; 500 ms frame swap while ≥1 active un-ignored Alert | — |
| R3 | Ignore mechanism via simple **INI** (codes, messages, notebooks, detectors) | ✅ Full | `config.ini` `[ignore]` + `[transient]`, hot-reloaded | — |
| R4 | Detect sync **failures** (incl. the silent case) | ✅ Full | Primary: local index ⟷ Graph timestamp divergence; corroborated by post-session error codes + OAlerts | detection latency = poll interval (default 5 min) |
| R5 | **Log every sync** and its **result** (success/failed) | ✅ Full (live, verified) | **ETW confirmed streaming every sync's `SendEvent` JSON in real time** (`detection-design.md`): notebook/section sync result, success/fail, error code, transient flag — seconds after it happens, OneNote open. Non-admin fallback (session-end log + polling) retained for machines without elevation. | requires the elevated collector running |
| R6 | Context menu on tray → **open the log / view history** | ✅ Revised per user (2026-09-06) | User asked for a minimal menu: *Show issues & status* / *Check now* / *Open config.ini* / *Quit*. The live status window carries sign-in/out, the cloud-check report and an *Open logs folder* button; all logs (sync history, collector, tray) are in one folder with 14-day retention (R16) | — |
| R17 | **Fail-closed detection**: only an explicit success counts as success; unclassified events treated as failures so nothing is missed | ✅ Full (reviewed 2026-09-06, 4 real gaps found & fixed) | Decision table in `SyncEventJson.Classify`; every `Storage.*` event parsed incl. unknown ones; other `OneNote.*` events picked up on a failure signal; field-name-independent success/error detection; malformed payloads and dropped ETW records surfaced; `[transient]` escalation implemented; 6 h TTL on unproven signals so they cannot pulse forever; `[ignore] events` to silence one once judged benign. Evidence + audit: `docs/fail-closed.md` | two name-signalled events (SyncBlockerInstantiated, GetUserTypesRequestFailed) show as warnings until judged benign |
| R18 | **Edit then close OneNote** must not lose a stranded change | ✅ Full | Verified OneNote emits no shutdown sync; collector re-ingests the just-unlocked session log on exit, tray runs an immediate cloud check, UploadStuck is Alert when OneNote is closed, and the cloud-check baseline persists across reboot. `docs/fail-closed.md` Part 1b | needs the one-time Graph sign-in; without it the icon is yellow (blind), never green |
| R19 | **Expired Microsoft token must alert** | ✅ Full | `GraphAuth.State` distinguishes NeverSignedIn (setup → Warn) from Expired (regression → **Alert**); MSAL `MsalUiRequiredException`/`MsalServiceException` set Expired. Verified end-to-end that the tray turns red on an alert-bearing status | — |
| R20 | **Two states only: error and success** — no middle 'warning' tier | ✅ Full | `SyncIssue` has no severity field; any un-ignored issue ⇒ red. `TrayIcons.State = {Ok, Error}`. Noise handled explicitly via `[ignore]`, `[transient]` hold-back, and 6 h TTLs. Tray starts red, goes green only after verifying. Verified end-to-end with the real binary (`icon Error → Ok`) | first cloud check has a bounded 2-min "not checked yet" window to avoid a red flash every login |
| R21 | **An error may only be cleared by a success that covers it**; otherwise it must persist | ✅ Full | `SyncScope` defines per-kind coverage (section-sync covers its section + real-time + its pages; notebook success covers only the notebook; page upload covers nothing; unknown events clear only by TTL) + success must postdate the failure + `_coveredUntil` stops a replayed stale failure re-opening a resolved error. 10 dedicated tests | grounded in OneNote's own schema, which reports section error state separately inside a notebook result |
| R16 | **All logs in one folder, detailed enough to troubleshoot, with a retention window** | ✅ Full | `C:\ProgramData\OneNoteWatcher\logs\`, daily files, `log_retention_days = 14` purge at start and nightly; collector logs ETW session lifecycle, every issue raise/recover, connectivity, backfill, 10-min summaries; tray logs icon transitions with reason, every Graph poll result, sign-in/out, dialogs, errors | token cache stays per-user (security) |
| R7 | Testing guide for provoking real errors | ✅ Full | `TESTING.md` (Phase 4) with 7 scenarios | — |
| R8 | No interaction with OneNote UI/process | ✅ Full | Read-only files + Graph + event log; no UIA, no COM (enforced by a unit test that greps the code for `OneNote.Application`) | — |
| R9 | Works on this corporate laptop (no admin, EDR) | ✅ Full | Every primary/secondary source is per-user, unprivileged | admin approaches are backup only |
| R10 | Survive monthly Office updates | ✅ Reasonable | depends on public Graph API + stable SQLite schema + JSON log format, not on offsets/UIDs; schema check degrades to a warning, never a crash | log/telemetry field names could shift; parser tolerates missing fields |
| R11 | Keep admin approaches as a **lower-priority backup** and explore them | ✅ Done | `detection-design.md`; matrix in `detection-design.md` now lists them as *Backup* not *Rejected* | — |
| R12 | Everything documented | ✅ Ongoing | this doc set | — |
| R13 | **Single install script / single uninstall script** (config, auto-start, unregister) | ✅ Full | `scripts/install.ps1` (build, copy, config, shared dir ACL, 2 Scheduled Tasks, start) · `scripts/uninstall.ps1` (stop, unregister, remove, `-PurgeData`) | run as admin |
| R14 | **Icon shows sync status whether or not OneNote is running** | ✅ Full | tray state derives from `status.json` (per-notebook last result + time) and local detectors, not from OneNote's presence; OneNote-not-running is only a tooltip note. No data at all → yellow, never blank | — |
| R15 | **Failure detail: which section/page, the reason (network? password? corruption?), and a recommended fix** — "a general error is useless" | ✅ Full (verified with real codes) | every failure → `SyncIssue` with WHERE (notebook / section / page resolved from the index, MRU cache and Graph), WHAT (`0xE000005D ErrFilePendingRename`), WHY (`ErrorCatalog` category: Network / Service / Permission / Storage / Password / Corruption / Conflict / FileState / ClientVersion / Transient) and FIX (concrete steps). Unknown codes still get a category from the description + the Microsoft article link. Connectivity events flag network causes explicitly. Shown in balloon, *Show issues* window and history line. | section names need the one-time Graph sign-in; without it the section shows as a short resource id plus the notebook |

## R5 in depth — "log every sync + result"

Met live, via ETW: every notebook/section sync result, page upload/download and connectivity change
is written to `logs/sync-history-<date>.log` within seconds, with the error code when one is present.
Periods when the collector was not running are back-filled from the finished session's diagnostic log
(OneNote's live log file is locked deny-read, so that part is necessarily post-session).

Without admin the collector cannot run at all; the tray then falls back to the cloud-side check and
dialog watching, and reports the missing collector as an error rather than showing green.
