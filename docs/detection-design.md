# How the watcher detects sync problems — and why

Every claim here was measured on the target machine (Windows 11, OneNote Win32 M365
16.0.20326.20112). Where measurement contradicted an assumption, the assumption lost — several
times, and those corrections are recorded because they are the most useful part of this document.

## The constraint that shaped everything

The watcher must run in the background and **never interact with OneNote's UI or process** — you are
working in that UI. It must also keep working with no admin rights on some machines. That rules out
UI Automation, driving dialogs, and COM (`OneNote.Application` also *launches* OneNote if it is
closed).

## What it uses

| Source | Gives | Needs |
|---|---|---|
| **ETW `OfficeLoggingLiblet`** — PRIMARY | every sync result live, seconds after it happens, with error code | admin (collector runs as SYSTEM) |
| **Office diagnostic log** — finished sessions | the same JSON, for periods the collector missed | nothing |
| **Microsoft Graph + local search index** | did my changes actually reach OneDrive? the only cloud-side proof | one-time sign-in |
| **`OAlerts` event log** | error dialogs OneNote shows | nothing |

### ETW is the primary source — verified, not assumed

`experiments/etw_probe.ps1` (run elevated) captured 60 s of `OfficeLoggingLiblet`
`{F50D9315-E17E-43C1-8370-3EDF6CC057BE}`; the `.etl` was decoded offline with `TraceEvent`.

- 38,359 of 43k events were one type, `etwtaskLogging`, with fields
  `wzProduct`, `wzCategory`, `wzTag`, `wzMessage`.
- **`wzMessage` carries the same `SendEvent {…}` JSON as the on-disk diagnostic log** — so one parser
  serves both, live and at rest.
- Latency: seconds.

A caveat that matters: the stream also contains *"Session has been sampled out, only critical events
will be sent to Aria"*. That governs **upload to Microsoft**, not local emission — the full
`SendEvent` fires locally regardless, so telemetry sampling does not blind the watcher.

### Why the cloud check exists

ETW tells you what OneNote *reported*. It cannot tell you whether content actually landed in
OneDrive when OneNote reported nothing at all — which is exactly what happens when you edit and then
close OneNote (see [fail-closed.md](fail-closed.md) Part 1b). Comparing the local search index with
Graph's `lastModifiedDateTime` is the only independent evidence, so it is worth the sign-in.

## Corrections — assumptions that measurement destroyed

These produced real false alarms in production. They are listed so nobody re-introduces them.

| Assumption | What the data showed |
|---|---|
| "Tail the live diagnostic log" (the original plan) | The running session's file is opened **deny-read**; `FileShare.ReadWrite` and `ReadWrite\|Delete` both throw. Readable only after the session ends. |
| "The index's `LastModifiedTime` is the last content edit" | It is **re-stamped when a notebook is opened or re-synced** — every *Work* section carried one 13-second burst while Graph showed edits months earlier. Comparing it absolutely to Graph produced **15 false "upload stuck" alerts**. The cloud check is now stateful: it baselines on first sight and judges only movement it observed. |
| "`SyncScore` reporting an error means a sync failed" | It is OneNote's **background replication scan**, listing what it *skipped* (e.g. a locked password-protected section). No notebook/section id, no clearing event — alerting on it is a permanent false alarm. Now recorded in history only. |
| "OneNote syncs when it closes" | It does **not**. Session logs simply stop; `SystemHealthUngracefulAppExitDesktop` shows exits are not always clean. An edit made just before closing can leave *no telemetry at all*. |
| "A notebook sync roughly every 30 min even when idle" | Measured gaps between **sync** events: median 0 m, p90 15 m, **max 114 m**. A 60-minute silence watchdog was guaranteed to false-alarm. Gaps between **any** Office telemetry never exceeded **17 m** — so liveness is watched there instead. |
| "A later success means the failure is over" | Only if it **covers** it. OneNote's `NotebookSyncResult` carries separate `IsSectionErrorSuppressed` / `IsSectionErrorUnexpected` fields, i.e. a notebook sync can succeed while its sections had errors. Coverage rules are in [fail-closed.md](fail-closed.md). |

## Approaches considered and rejected

| Approach | Verdict |
|---|---|
| UI Automation of the sync dialog | **Rejected by requirement** — it would fight you for the UI. (Also measured: OneNote's UIA tree exposes ribbon chrome only.) |
| COM `IApplication` | No sync/error attribute exists on any node (`GetHierarchy` gives `name, ID, path, lastModifiedTime, color, isCurrentlyViewed, …`). `SyncHierarchy` only *triggers* a sync. Instantiating the COM object launches OneNote. |
| `%TEMP%\OneNote.log` via the `Logging` registry key | The only switch in `onmain.dll` is `EnableObjectModelLogging` — COM call logging, not sync. |
| Parsing the `.onecache` binary | Proprietary; shows activity, not outcome. Superseded by the index + Graph comparison. |
| OTele SQLite store (`OTele\onenote.exe.db`) | Contains the same events Bond-encoded, but only flushed at exit/upload failure — 0 rows across 115 s of forced syncs. Kept as a documented last resort; `experiments/bond_decoder.py` decodes it. |
| DLL injection / API hooking | **Rejected outright.** No Office symbols, offsets move monthly, EDR present, and it risks corrupting the app being watched. |
| Network interception (TLS proxy) | Invasive, needs cert trust, cannot map a failure to a notebook. |

## Privilege notes

Starting an ETW session requires elevation — `logman … -ets` returns `Access is denied` for a normal
user (verified). The collector therefore runs as SYSTEM via a Scheduled Task. If elevation is
unavailable on some machine, the tray degrades to the cloud check plus dialog watching, and reports
the missing collector as an error rather than showing green.

## Reproducing the evidence

`experiments/` holds the read-only probes behind every measurement above, and
`OneNoteWatcher.Collector.exe --audit` re-runs the event classifier over your real diagnostic logs.
See [reference-data-sources.md](reference-data-sources.md) for exact paths and formats.
