# Local Machine Findings — Verified on Target (2026-09-04)

Empirical verification of the assumptions in `onenote-sync-issue-monitoring-research.md`,
run against the actual target machine. **Where this document and the research doc disagree,
this document wins** — the research is desk research, this is measured.

## Environment

| Item | Value |
|---|---|
| OS | Windows 11 Enterprise 10.0.26200 |
| OneNote | `C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE`, **16.0.20326.20112** (Win32 M365 — correct target) |
| .NET SDKs | 8.0.404, 10.0.302 |
| Diagnostics dir | `%LOCALAPPDATA%\Temp\Diagnostics\` — exists, with `EXCEL`, `ONENOTE`, `OUTLOOK` subfolders |

## Finding 1 — The live diagnostic log is LOCKED deny-read (BLOCKER for real-time tailing)

The log file belonging to the **running** OneNote process cannot be opened, even with
`FileShare.ReadWrite`. Rotated (previous-session) files open fine.

```
OK       2026-09-02 13:21:34  Primary1788330094610356700_....log
OK       2026-09-03 07:29:22  Primary1788395362908713500_....log
OK       2026-09-04 07:47:32  Primary1788482852896666000_....log
OK       2026-09-04 15:55:30  Primary1788512130428618100_....log
LOCKED   2026-09-04 22:54:17  Primary1788537257202678300_....log   <- active session
LOCKED   2026-09-04 22:54:17  Primary1788537257201926200_....log   <- active session
```

`[System.IO.File]::Open(path,'Open','Read','ReadWrite')` throws
`IOException: because it is being used by another process` on the active files.

**Consequence:** the research's "Stage 1 — FileSystemWatcher + tail the newest file" cannot
see the current session's events at all. Content becomes readable only after the file is
released (OneNote exits / rotates). For a long-running OneNote that is hours-to-days of
detection latency. **Log tailing cannot be the primary real-time signal.**

## Finding 2 — No sync or error events are being logged at all

Across ~45 KB of *all* real content in every OneNote log on this machine (3 days, 6 sessions):

| Pattern | Occurrences |
|---|---|
| `Error.Code` | 0 |
| `Error.Type` | 0 |
| `jerrc` | 0 |
| `OneStore` | 0 |
| `FSSHTTP` | 0 |
| `Notebook` | 0 |
| `Sync` / `sync` | 1 (unrelated) |

The `0xE000xxxx` / `0xE40xxxxx` codes the research is built around: **zero matches**.

What *is* logged is only generic Office bootstrap telemetry, all emitted in the first seconds
of app start: `Office.Licensing.*`, `Office.Telemetry.*`, `Office.System.SystemHealth*`,
`Office.Text.*`, `Office.Experimentation.*`, `Office.Identity.*`, `Office.Manageability.*`.

The `Additional\` log stream (a second stream the research does not mention) is **empty** —
71 bytes, header row only:
`Timestamp  Process  TID  Area  Category  EventID  Level  Message  Correlation`

Each 16 MB `.log` file is preallocated and null-padded; real content is only 10–40 KB.
**Parsers must strip `\0` padding.**

`%LOCALAPPDATA%\Temp\Diagnostics\UploadCache` — **does not exist** on this machine.

## Finding 3 — Not explained by a disabled privacy setting

All relevant policy/telemetry keys are absent or empty, i.e. Office is at its **default**
diagnostic level, not a locked-down enterprise one:

- `HKCU:\Software\Microsoft\Office\common\privacy` — absent
- `HKCU:\Software\Microsoft\Office\16.0\common\privacy` — present, no values
- `HKCU:\Software\Microsoft\Office\16.0\common\clienttelemetry` — present, no values
- `HKLM:\SOFTWARE\Policies\Microsoft\office\...\privacy` — absent

So this is the out-of-the-box behaviour, not a misconfiguration to flip. Raising the Office
diagnostic-data level to "Optional" is worth testing, but must not be assumed to work.

## Finding 4 — Cache directory is present and structured as documented

`%LOCALAPPDATA%\Microsoft\OneNote\16.0\` contains `cache`, `Backup`, `MasterIndex`,
`ServerListings`, `FullTextSearchIndex`, `AccessibilityCheckerIndex`. Viable for the
staleness/backlog heuristic.

## Impact on the plan

The research ranked **log tailing #1 (primary) and UIA #2 (fallback)**. On this machine that
ordering is **inverted** — findings 1 and 2 are each independently fatal to log tailing as a
primary real-time signal.

**Therefore: UIA is the primary detector; the log detector is opportunistic; cache staleness
is the backstop.** The app must treat detectors as pluggable and degrade gracefully when a
detector reports "unavailable", rather than assuming any single signal works.

## Still unverified (needs a run against real OneNote on this box)

- The exact UIA tree of the "Shared Notebook Synchronization" dialog on build 16.0.20326.20112
  (AutomationIds, ControlTypes, row structure, where the error code/message text lives).
- Whether the nav-pane per-notebook sync glyph is exposed as a readable UIA element or is
  drawn (in which case UIA cannot read it).
- Whether raising Office diagnostic data to "Optional" makes sync errors appear in the logs.
- Whether a real sync error changes anything observable under `cache\`.
