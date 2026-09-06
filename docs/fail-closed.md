# Sync-Failure Coverage — the fail-closed contract

**The rule.** A sync failure is rare and costly to miss, so: *only an explicit success signal counts
as success; anything the watcher cannot prove successful must surface.* This applies to **the whole
system**, not just event classification — a watcher that cannot see must never look healthy. Extra
alerts are cheap; a missed failure defeats the app's purpose.

Everything below was verified against real data on the target machine; 140 tests guard it, including
a dedicated fail-closed suite and coverage suite.

## Two states, on purpose

The watcher shows **error** or **success** — nothing in between. `SyncIssue` has no severity field at
all: if an un-ignored issue exists, the tray is red. A middle "warning" tier is the tier people learn
to ignore, which is exactly how a rare sync failure gets missed.

Noise is therefore controlled *explicitly*, never by softening the signal:

| Mechanism | What it does |
|---|---|
| `[ignore]` (codes, messages, notebooks, sections, detectors, **events**) | you decide something is not worth seeing |
| `[transient]` | you list codes that are held back until they repeat N times in a window — still written to the history |
| `ExpiresUtc` (6 h) | an unproven signal self-clears, so it cannot stay red forever without evidence |

The tray starts in **error** and only turns green once it has verified — never the other way round.
The one bounded exception is the very first cloud check: for up to 2 minutes after launch, "not
checked yet" is not reported as blindness (otherwise every login would flash red and train you to
ignore red). After that bound, an unresolved check *is* an error, so nothing can hide there.

> **Terminology note.** In this repo the catastrophic case — *not alerting when a sync really failed*
> — is written out in words rather than as "false positive/negative", because those labels are easy
> to swap. The rule is simply: **never silently miss a failure; err toward telling the user.**

## Part 1 — the system-wide rule: silence is not success

Classifying events correctly is useless if the pipeline delivering them is dead. These checks run on
a timer in both processes (`Core/Health/HealthIssues.cs`), and every one produces a visible issue
with an explanation and a fix:

| Condition | Why it matters |
|---|---|
| **OneNote running, but NO Office telemetry at all for `pipeline_timeout_minutes` (default 45)** | The single most dangerous state: a dead ETW pipeline is otherwise indistinguishable from healthy sync. Keyed on *any* telemetry, not sync events — see the measurement below. Clamped to the collector's own uptime, and suppressed while OneNote is closed. |
| Collector not running / status file stale > 2 min | Real-time detection is off |
| Local OneNote index unreadable (missing, locked, partial) | Names and the cloud comparison are incomplete |
| **Microsoft sign-in EXPIRED** (an account exists but can no longer get a token) | The cloud check is the *only* thing that can prove edits reached OneDrive — especially after OneNote closes (Part 1b). Losing it is a silent regression from a working state. (Never having signed in is setup rather than a regression, but it is still an error — the watcher is blind either way.) |
| Graph failing, or last success > 90 min | The cloud-side check is blind |
| `OAlerts` subscription unavailable | OneNote's error dialogs are not being watched |
| History/log writes failing | The record of sync results is being lost |
| A detector switched off in `config.ini` | Monitoring silently reduced |
| `config.ini` unreadable → defaults | **Ignore rules would not be applied** |
| ETW records dropped by Windows | A sync event may not have been seen |
| Unhandled exception in either process | Logged, and the collector is marked not-running — a crash leaves evidence, never a frozen green icon |

The tray also guards its own refresh loop: an internal error turns the icon **red** with "internal
error, see tray log" rather than freezing on a stale state.

### Why the watchdog watches telemetry, not sync events

The first production run of this watchdog produced a **false alarm**, and fixing it needed real data
rather than an assumption. Measured across every readable session log on the target machine:

| Gap between… | median | p90 | **max** | gaps > 60 min |
|---|---|---|---|---|
| OneNote **sync** events | 0.0 min | 14.6 min | **114.3 min** | 2 |
| **Any** Office telemetry | 0.0 min | 0.5 min | **17.3 min** | 0 |

An idle OneNote legitimately goes ~2 hours without a sync event, so "no sync for an hour" is normal
and alerting on it fires every time you leave OneNote open. Telemetry, by contrast, never stopped for
more than 17 minutes — so silence *there* really does mean the pipeline died. The watchdog therefore
keys on telemetry liveness, with a 45-minute default (2.6× the observed maximum).

Two bugs were fixed together here:

1. **The watchdog counted time before the collector was running.** After a restart the post-exit
   backfill sets "last activity" from *historical* log data; the watchdog compared that to the clock
   and reported 2.6 h of silence **5 seconds after starting**. Silence is now clamped to the
   collector's own uptime — it can never claim to have watched a period it did not.
2. **The threshold rested on a wrong assumption** ("a sync roughly every 30 min even when idle"),
   which the project's own captured data had already contradicted.

`status.json` now carries `LastTelemetryUtc` alongside `LastSyncEventUtc`, and the issues window shows
both — so "quiet but alive" is distinguishable from "dead" at a glance.

### How the dead-pipeline path was actually tested

Being precise, because "it's covered" is worth little without saying how it was checked:

| Layer | How it was verified |
|---|---|
| Decision logic | Unit tests with an injected clock and a stubbed process check: alert after the timeout while OneNote runs; **no** alert while OneNote is closed; clears when activity resumes |
| Collector → tray seam | Integration test: real health sweep → `Snapshot()` → **real `status.json` on disk** → `WatcherStatus.Load` reads the Alert back with its summary and fix |
| Tray behaviour | **End-to-end with the real built binaries**, pointed at an isolated data folder: (a) a stale `status.json` claiming `CollectorRunning=true` → `icon Error … [errors=1 :: collector is not running]`; (b) a `status.json` carrying the stale-activity issue → `icon Error … no sync telemetry for 2.0 h` (red, pulsing); (c) a healthy `status.json` → `icon Error → Ok: in sync [errors=0]` (the tray starts red and only goes green once it has actually verified) |

**Not exercised:** stopping the *live ETW session itself*, because creating one needs admin and the
installed collector runs as SYSTEM. That seam is exactly what the timeout is designed to catch
without needing to detect the session's death, and everything downstream of it is covered above.

This work also fixed a real observability bug: the tray only logged on a *colour* change, so a tray
that started in the state it computed left no evidence of its decision. It now logs whenever the
reason **or the issue counts** change — which is how the 2 → 1 transition above became visible.

## Part 1b — the "edit, then close OneNote" scenario

**Measured on real sessions: OneNote emits NO final sync when it closes.** Session logs simply stop —
the last sync events land whenever they happen, and `SystemHealthUngracefulAppExitDesktop` shows exits
are not always clean:

```
16:50:44  NotebookSyncResult Success=true
18:45:01  PageSyncSession            ← ~2 h later, then the session just ends
```

So an edit made shortly before closing can produce **no telemetry at all**. Since the watchdog is
suppressed while OneNote is closed, this is exactly where a stranded change could hide. What covers it:

| Step | Mechanism |
|---|---|
| OneNote closes | The collector detects the transition, writes `(OneNote) CLOSED` to the history, and **re-ingests the session log that has just unlocked** — anything the live ETW session missed is still captured |
| Immediately after close | The tray triggers an **immediate cloud check** rather than waiting for the next poll |
| Stranded changes found | `UploadStuck` — an error like any other. With OneNote closed nothing will retry, so it is *more* serious, not less |
| PC is shut down too | The cloud-check baseline is **persisted** (`outcome-state.json`). Without it, the next login would re-baseline the stranded change as normal and lose it. A test asserts a fresh detector misses it and the restored one does not |

**Honest limitation:** this path depends on the Microsoft Graph check, so it needs the one-time sign-in.
If you are not signed in, the watcher cannot verify that anything reached the cloud — and it says so
(`HealthIssues.GraphBlind`, red icon) rather than showing green.

## Part 2 — event classification

### The decision table

Implemented in `SyncEventJson.Classify`. **Order matters** — it is evaluated top to bottom:

| # | Condition | Outcome | Result |
|---|---|---|---|
| 1 | `ErrorState_TimeIn*SyncErrorState > 0` — a page really spent time failing | `SuspectedFailure` | error (expires after 6 h) |
| 2 | Event is a documented non-outcome (see exceptions below) | `Diagnostic` | history only |
| 3 | `Success`/`IsSuccess`/`Succeeded`/`WasSuccessful` **= false** | `Failure` / `Transient` | error |
| 4 | Any `*Error_Code`/`*ErrorCode` ≠ 0, or any error-text field with a real message | `Failure` / `Transient` | error |
| 5 | Event **name** contains Fail/Error/Blocker/Stuck/Corrupt/ReadOnly/Conflict/Denied/Unauthori/Expired/Rejected/Abort/Crash/Inconsistenc | `SuspectedFailure` | error (expires after 6 h) |
| 6 | Success flag **= true** | `Success` | clears the issue |
| 7 | An error field explicitly says `"No error"` | `Success` | clears the issue |
| 8 | A page upload/download completed (only emitted on completion) | `Success` | clears the issue |
| 9 | **anything else** | `Unknown` | error (expires after 6 h) |

Row 9 is the fail-closed default: a `Office.OneNote.Storage.*` event this build has never seen, or a
known event that arrives *without* its success flag, is treated as a possible failure — not as "fine".

### Scope

- **Every `Office.OneNote.Storage.*` event is parsed**, known or not.
- Other `Office.OneNote.*` events are picked up **when they carry a failure signal** in the name or
  fields, so a future `…Failed` / `…Blocked` event is caught with no code change. Benign app noise
  (AppLaunch, Copilot licence, ConfigServiceReady) stays out.
- Non-OneNote Office events are out of scope.

### Field-name independence

Success is read from `Data.Success`, `Data.IsSuccess`, `Data.Succeeded`, `Data.WasSuccessful`.
Error codes from **any** field matching `Data.*Error_Code` / `Data.*ErrorCode` (so `SH_ErrorCode`,
`NotebookErrorCode`, `Error_Code` and future spellings all count). Error text from `Data.Error`,
`Error_Description`, `Error_Type`, `OperationWithError`, `FailureReason`, `ErrorMessage`.
A value of `""`, `"No error"`, `"None"`, `"OK"`, `"Success"`, `"0"` is treated as *not* an error.

### The two documented exceptions

Both are deliberate, evidence-backed, and recorded in the history (never silently dropped):

| Event | Why it is not an outcome |
|---|---|
| `Storage.SyncScore` | OneNote's **background replication scan**. It reports items it *skipped* (e.g. a locked password-protected section) at every session start and periodic pass. It carries no notebook/section id and has no clearing event, so alerting on it is a permanent false alarm — this was a real false alert on 2026-09-06. |
| `Storage.PageSyncSession` | Per-page timing metric. Its `ErrorState_Time*` fields *are* used — row 1 above catches a page that actually spent time in an error state. |

`Storage.ConnectivityChanged` is not an outcome either; it drives a dedicated offline warning.

## When an error clears — coverage rules

Clearing is as strict as raising. **An error is removed only when a later success genuinely COVERS
it**; a success that proves nothing about the failed content leaves the error standing.

| A success of this kind… | …covers | …does NOT cover |
|---|---|---|
| **Section sync** for section X | X's sync failures, X's real-time failures, page failures known to be in X | anything in another section |
| **Notebook sync** | notebook-level failures | **its sections** — OneNote's own `NotebookSyncResult` carries separate `IsSectionErrorSuppressed` / `IsSectionErrorUnexpected` fields, i.e. a notebook sync can succeed *while its sections had errors*, so it is not evidence a section recovered |
| **Real-time session** ("No error") for section X | X's real-time channel failures | X's full section-sync failure (different mechanism, weaker evidence) |
| **Page upload / download** | nothing | anything — one page reaching the server says nothing about the rest of the section |
| anything | — | an **unclassified/suspected** event: we cannot know what would prove it resolved, so it clears only by its 6 h TTL |

Two ordering rules on top of that:

- A success must be **later than the failure** it clears. The post-exit backfill replays finished
  session logs, so out-of-order arrival is real, not theoretical.
- Symmetrically, a **stale failure replayed after a covering success does not re-open** the error
  (`_coveredUntil` per scope) — otherwise re-ingesting a session log at OneNote exit would resurrect
  errors that were already resolved.

**Two real bugs this fixed.** The previous build keyed page uploads and section syncs to the *same*
scope, so a single page upload cleared a whole section's failure; and it had no timestamp check, so
an older success could clear a newer failure. Both are now covered by tests that fail on the old logic.

Recovery is expected and encouraged: the tray starts in **error** and goes green as soon as it has
actually verified — that is the intended lifecycle, not a bug. What must never happen is going green
without that proof.

## When an issue clears

| Outcome | Clears when |
|---|---|
| `Failure` (non-transient code) | a proven success for the same scope |
| `Failure` (code catalogued as transient, or the real-time channel) | a proven success for the same scope |
| `Transient` (OneNote flagged it retryable) | a proven success |
| `SuspectedFailure` / `Unknown` | a proven success, **or 6 h TTL** |

The TTL exists because an unproven signal may have no clearing event; without it a single odd event
would stay red forever. A **confirmed** failure has no TTL — it clears only on a real success.

Only a *proven* success clears an issue, and it must be for the **same scope** (section / notebook /
event name). A success elsewhere never masks a failure here.

## Coverage beyond event classification

| Silent-miss risk | Mitigation |
|---|---|
| A payload that cannot be parsed | Returned as `Unknown` (not dropped), counted in `MalformedPayloads`, logged |
| Windows drops ETW records under load | `EventsLost` surfaced as an error telling you to check OneNote's own sync status |
| ETW session dies | Auto-restart with backoff; tray shows "collector not running" |
| Collector not running at all | Tray icon goes red; the cloud-side Graph check still runs |
| Events emitted while the collector was down | Diagnostic-log backfill of finished sessions at startup |
| A failure with no telemetry at all | Independent cloud-side check (local index ⟷ Graph) + OAlerts dialog watcher |

## Gaps found and fixed by this review

1. **Unknown events were dropped.** The parser had a hard-coded list of eight event names; anything
   else — including a future failure event — was discarded. Now every `Storage.*` event is parsed.
2. **Three real events were being ignored**, confirmed present in this machine's logs:
   `Storage.RealTime.SyncBlockerInstantiated`, `UserInfoService.GetUserTypesRequestFailed`, and error
   fields on `GetSharePointIdsForDocument` (`Data.IsSuccess`) / `Navigation.Navigate` (`Data.SH_ErrorCode`).
3. **A missing success flag defaulted to "OK".** Now `Unknown`.
4. **The diagnostic log's trailing Correlation column broke JSON parsing** — the row is `{…}\t<guid>`,
   so `JsonDocument.Parse` failed with "extra data" and ~4 events per session were silently lost from
   the backfill. Fixed by reading exactly one JSON value. The audit's malformed count went 2 → **0**.
5. **`[transient]` was advertised in config.ini but never implemented.** Now implemented with
   escalation.

## Evidence — audit over this machine's real logs

`OneNoteWatcher.Collector.exe --audit` runs every OneNote event in the Office diagnostic logs through
the **real classifier** (no admin needed) and reports the result. Run on 2026-09-06:

```
SendEvent rows: 1257;  Office.OneNote.* rows: 113;  malformed OneNote payloads: 0

IN SCOPE — every event below was classified (a 'Success' requires positive proof):
  Storage.ConnectivityChanged                      6  Diagnostic=6
  Storage.NotebookSyncResult                      24  Success=24
  Storage.PageSyncSession                          6  Diagnostic=6
  Storage.RealTime.NoteItHttpDownload              4  Success=4
  Storage.RealTime.NoteItHttpUpload                8  Success=8
  Storage.RealTime.NoteItService                  22  Success=22
  Storage.RealTime.SyncBlockerInstantiated         2  SuspectedFailure=2
  Storage.SectionSyncResult                       17  Success=17
  Storage.SyncScore                                2  Diagnostic=2
  UserInfoService.GetUserTypesRequestFailed        2  SuspectedFailure=2

DELIBERATELY OUT OF SCOPE (no failure signal in name or fields):
  Office.OneNote.Navigation.Navigate               6
  Office.OneNote.System.AppLifeCycle.AppLaunch     2
  Office.OneNote.System.ConfigServiceReady         2
  Office.OneNote.Augmentation.Copilot.DisabledNoLicense 2
  Office.OneNote.GetSharePointIdsForDocument       1
```

Every `Success` above came from an explicit flag, an explicit `"No error"`, or a completed transfer —
never from an absence of evidence.

## Known noise, and how to silence it

The two `SuspectedFailure` events above are name-signalled only; whether they are benign is unproven,
so fail-closed keeps them visible as errors that self-expire after 6 h.
Once you are satisfied they are harmless:

```ini
[ignore]
events = Office.OneNote.Storage.RealTime.SyncBlockerInstantiated, Office.OneNote.UserInfoService.GetUserTypesRequestFailed
```

They stay in the sync history either way. This is the deliberate trade: the watcher errs toward
telling you, and you decide what to mute — rather than the watcher deciding silently for you.

## Tests guarding the contract

`tests/OneNoteWatcher.Core.Tests/FailClosedTests.cs` and the collector's detector tests cover:
unknown Storage events → `Unknown` + warning; unknown event *with* an error code → Alert; failure
words in a name; malformed payloads; every alternative success/error field spelling; missing success
flag; out-of-scope noise staying out; transient escalation and window expiry; SyncScore staying
diagnostic; TTL expiry; only a same-scope proven success clearing an issue; dropped ETW records;
and a regression guard listing every event name observed on this machine.
