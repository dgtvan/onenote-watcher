# onenote-watcher

A background tray app that tells you when OneNote stops syncing — and tells you **where** (notebook /
section / page), **what** (error code and symbolic name), **why** (category) and **how to fix it**.

It never touches OneNote's UI or process: it reads OneNote's own telemetry live over ETW, its local
index, the Windows event log, and Microsoft Graph.

Targets the Win32 Microsoft 365 build of OneNote (`ONENOTE.EXE`) on Windows 10/11.

## Install / uninstall (Administrator PowerShell)

```powershell
.\scripts\install.ps1                 # full install: build, register auto-start, start (replaces config.ini)
.\scripts\install.ps1 -UpdateOnly     # after a code change: swap binaries only (keeps config.ini)
.\scripts\uninstall.ps1               # remove app + auto-start (keeps logs; -PurgeData removes them)
```

A **full install replaces `config.ini`** with the copy from this repo, so a reinstall gives a clean,
predictable state and newly added settings actually arrive. Your previous file is saved beside it as
`config.ini.bak-<timestamp>`. Use **`-UpdateOnly`** when you have settings to keep: it swaps the
binaries and never touches your config.

Then once, from the tray: **Show issues & status → Sign in to Graph…** (a device code; gives section
names in failures and the cloud-side check). Your Graph client id is already in
[config.ini](config.ini) — see [docs/setup-graph.md](docs/setup-graph.md).

Everything lives in one folder, `C:\ProgramData\OneNoteWatcher`: both executables, `config.ini`,
`status.json` and `logs\`.

## Two states, on purpose

| Icon | Meaning |
|---|---|
| **red cloud, pulsing** | **ERROR** — something is wrong and you need to look |
| white-outlined cloud | **SUCCESS** — everything the watcher can check is healthy |

There is no middle "warning" tier, because that is the tier people learn to ignore. Anything the
watcher cannot verify counts as an error — **including its own blindness** (collector down, Microsoft
sign-in expired, index unreadable, telemetry pipeline dead). Noise is controlled explicitly instead,
through `[ignore]` rules, the `[transient]` hold-back and time-limited unproven signals.

An error clears only when a later success genuinely **covers** it. A success that proves nothing
about the failed content leaves the error standing. See [docs/fail-closed.md](docs/fail-closed.md).

## Tray menu

**Show issues & status** (live window; Graph sign-in and *Open logs folder* live there) ·
**Check now** · **Open config.ini** · **Quit**

## Logs — one folder, 14-day retention

`C:\ProgramData\OneNoteWatcher\logs\`, one file per day:

- `sync-history-<date>.log` — every sync result, success or failure
- `collector-<date>.log` — ETW session, issues raised/cleared, backfill, health
- `tray-<date>.log` — icon state changes with the reason, Graph polls, dialogs

## Development

```powershell
dotnet build ; dotnet test                                    # 140 tests
.\src\OneNoteWatcher\bin\Debug\net8.0-windows\OneNoteWatcher.exe --simulate
OneNoteWatcher.Collector.exe --audit                          # classifier coverage over real logs
OneNoteWatcher.exe --graph-check                              # local index vs Microsoft Graph
```

## Docs

| Doc | What it is |
|---|---|
| [architecture.md](docs/architecture.md) | The two processes, the shared folder, the config reference |
| [detection-design.md](docs/detection-design.md) | How detection works, what was measured, and the assumptions that measurement destroyed |
| [fail-closed.md](docs/fail-closed.md) | The contract: only proven success counts, and when an error is allowed to clear |
| [reference-data-sources.md](docs/reference-data-sources.md) | Exact paths, schemas, event fields, epochs |
| [setup-graph.md](docs/setup-graph.md) | One-time Entra app setup |
| [testing.md](docs/testing.md) | How to provoke real sync failures and watch the app react |
| [requirements.md](docs/requirements.md) | Every stated requirement and an honest verdict |
| [experiments/](experiments/) | The read-only probes behind the measurements |

## Credits

Tray icon: [Cloud icon](https://www.flaticon.com/free-icon/cloud_2311548?related_id=2311548) by
[Flaticon](https://www.flaticon.com/), used under the Flaticon Free License with attribution. The app
recolours it at runtime (white outline for success, red for error) — see
[TrayIcons.cs](src/OneNoteWatcher/TrayIcons.cs).
