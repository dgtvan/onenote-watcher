# Architecture

Two processes, one folder, one config file.

```
┌── ELEVATED collector ────────────────────────┐   ┌── UNELEVATED tray (your session) ────────┐
│ SYSTEM · Scheduled Task at startup           │   │ you · Scheduled Task at logon            │
│                                              │   │                                          │
│ EtwSyncDetector      live ETW, seconds       │   │ TrayController   two-state icon          │
│   OfficeLoggingLiblet → wzMessage            │   │ OAlertsDetector  OneNote error dialogs   │
│   → SendEvent{json} → SyncEventJson          │   │ GraphPoller      cloud-side check        │
│ DiagLogBackfill      finished sessions       │──▶│ IssuesForm       live status window      │
│ health sweep         pipeline liveness       │ f │ health checks    can we even see?        │
│                                              │ i │                                          │
│ writes status.json + logs/                   │ l │ reads status.json, writes section-names  │
└──────────────────────────────────────────────┘ e └──────────────────────────────────────────┘
        shared parser: SyncEventJson — one JSON shape for live ETW and the on-disk log
```

The collector needs elevation only because creating an ETW session does. The tray must stay
unelevated so it can own a tray icon in your session — hence the split, and hence
`C:\ProgramData` rather than `Program Files` (the unelevated side must be able to write).

If the collector is absent (no admin on some machine), the tray still runs: cloud check + dialog
watching, with the missing collector itself reported as an error.

## Projects

| Project | Target | Contains |
|---|---|---|
| `OneNoteWatcher.Core` | `net8.0` | event classifier, error catalog, name resolution, ignore rules, INI, history, status model, health issues. No Windows UI — fully unit-testable |
| `OneNoteWatcher.Collector` | `net8.0-windows` | the only privileged code: ETW session, `EtwSyncDetector`, coverage rules, diagnostic-log backfill, `--audit` |
| `OneNoteWatcher` | `net8.0-windows` | WinForms tray, issues window, Graph auth/poller, OAlerts |
| `*.Tests` | xUnit | 140 tests, incl. fixtures captured from this machine |

## The one root folder — `C:\ProgramData\OneNoteWatcher`

| File | Writer | Content |
|---|---|---|
| both `.exe` + `config.ini` | installer | the app itself |
| `status.json` | collector | collector/OneNote alive, last telemetry, last sync, per-notebook state, active errors |
| `section-names.json` | tray | Graph section-id → "Notebook / Section", so the collector can name sections |
| `outcome-state.json` | tray | cloud-check baseline, persisted so a reboot cannot lose a stranded change |
| `processed-sessions.txt` | collector | diagnostic-log sessions already ingested |
| `logs/*.log` | both | one file per day, purged after `log_retention_days` |

## Data flow for one sync

1. OneNote emits `Office.OneNote.Storage.*` telemetry.
2. The collector's ETW session receives it within seconds and `SyncEventJson` classifies it
   ([fail-closed.md](fail-closed.md)).
3. Every event is written to `logs/sync-history-<date>.log`; problems become errors in `status.json`.
4. The tray reads `status.json` every 3 s, applies `[ignore]`, and drives the icon.

## Configuration

All in [config.ini](../config.ini), next to the executables. Key settings:

| Key | Default | Meaning |
|---|---|---|
| `general.poll_minutes` | 5 | how often the cloud-side check runs |
| `general.grace_minutes` | 10 | how long a local change may sit un-uploaded before it counts as stuck |
| `general.pipeline_timeout_minutes` | 45 | no Office telemetry at all for this long (while OneNote runs) ⇒ dead pipeline |
| `general.log_retention_days` | 14 | logs older than this are deleted |
| `graph.client_id` / `tenant` | — | Entra app, see [setup-graph.md](setup-graph.md) |
| `etw.shared_dir` | the exe's folder | the one root folder |
| `[ignore]` | — | suppress by code, message, notebook, section, detector or **event name** |
| `[transient]` | empty | codes held back until they repeat N times — the only deliberate hold-back |

## Design rules that are not negotiable

- **Never touch OneNote's UI or process.** No UI Automation, no `OneNote.Application` COM (it would
  also *launch* OneNote). Everything is read-only files, ETW, the event log, and Graph.
- **Two states only**, error and success — see [fail-closed.md](fail-closed.md).
- **Only proven success counts as success**, and an error clears only on a success that covers it.
