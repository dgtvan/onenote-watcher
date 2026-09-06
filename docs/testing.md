# Testing Guide — Provoking Real OneNote Sync Problems

How to verify the watcher reacts to genuine sync failures. Scenarios are ordered safest-first and
use a throwaway notebook so your real notes are never at risk. Each lists steps, what to expect, and
how to undo.

## 0. Install (one script) and sign in once

```powershell
# Admin PowerShell — builds Release, installs to C:\Program Files\OneNoteWatcher, registers both
# auto-start tasks (collector as SYSTEM at startup, tray at your logon), starts them:
& "D:\Src\Personal\onenote-watcher\scripts\install.ps1"

# later, to remove everything (add -PurgeData to also delete the history):
& "D:\Src\Personal\onenote-watcher\scripts\uninstall.ps1"
```

Then tray menu → **Sign in to Graph…** — a dialog shows a code (copied to your clipboard) and opens
the Microsoft device-login page; approve "Read your OneNote notebooks". One time only.

Dev loop without installing: `dotnet build`, then
`.\src\OneNoteWatcher\bin\Debug\net8.0-windows\OneNoteWatcher.exe` (the collector needs admin:
`.\src\OneNoteWatcher.Collector\bin\Debug\net8.0-windows\OneNoteWatcher.Collector.exe` from an admin shell).

**Show issues & status** opens the live WHERE / WHAT / WHY / FIX window (sign-in, cloud-check report
and *Open logs folder* buttons are at its top). All logs: `C:\ProgramData\OneNoteWatcher\logs\`
(`sync-history-<date>.log` is the per-sync history), kept 14 days.

**Expected icon when everything is installed and healthy:** green, tooltip "in sync — last sync
HH:mm". Quit OneNote → still green with "(OneNote not running)". Stop the collector task → yellow
"collector not running; last known state OK at HH:mm". The icon is never blank.

## 1. Simulated alert (no OneNote needed) — proves the icon + menu

```powershell
.\src\OneNoteWatcher\bin\Release\net8.0-windows\OneNoteWatcher.exe --simulate
```

**Expect:** tray icon turns **red and pulses**; a balloon "OneNote sync problem — Watcher Test /
Demo Section 0xE000005D"; **Show issues** lists the simulated alert; adding `codes = 0xE000005D` to
`config.ini [ignore]` and relaunching suppresses the red (issue still shows as "(ignored)").
**Undo:** quit from the menu.

## 2. Upload stuck — the real silent-failure case (no admin needed)

1. Create a throwaway notebook **Watcher Test** on your OneDrive; add a page.
2. Turn **Wi-Fi off** (Windows quick settings — no admin needed).
3. Edit the page (type a line). OneNote saves locally; upload cannot happen.
4. Wait `grace_minutes` + one poll (defaults: ~10–15 min; set `grace_minutes = 2`,
   `poll_minutes = 1` in `config.ini` to see it in ~3 min).

**Expect (collector, near-real-time):** OneNote retries the section sync while offline; the history
shows the results and an `Offline` warning (yellow) from the connectivity event; on reconnect, the
section result goes green. **Expect (baseline cloud check, no collector):** the section's local stamp
moved while watching and the server did not follow → after `grace` an `UploadStuck` alert naming
*Watcher Test / <section>*; red pulsing icon. (The baseline never alerts on pre-existing skew at
startup — it only judges changes it observed.)
**Expect (collector installed):** additionally, when connectivity returns and the sync retries, a
live history line for the section result.
**Undo:** turn Wi-Fi back on → next poll clears the issue; the history shows the recovery.

## 3. Download stuck (needs a second device)

With Wi-Fi **off** on this PC, edit the same page from the OneNote web app / phone. Reconnect is
not done yet on this PC → after `grace` the watcher reports the notebook behind the server
(`DownloadStuck`). **Undo:** reconnect; issue clears.

## 4. Error-dialog path (real, immediate) — OAlerts

Open a notebook link that no longer exists (OneNote shows *"We couldn't open that location…"*).
**Expect:** an `ErrorDialog` issue within seconds (OAlerts detector). Add
`messages = *couldn't open*` to `[ignore]` to confirm suppression. **Undo:** dismiss the dialog.

## 5. Error-code path (the exact code) — collector, near-real-time

Prereq: collector installed and running (scenario 0).

1. In *Watcher Test*, password-protect a section (Review → Password → Set Password), then lock it
   (right-click the tab → Lock).
2. Force a sync (Shift+F9). OneNote emits `SectionSyncResult` / `SyncScore`.
3. **Expect:** within seconds, a `sync-history.log` line and — if the section errors —
   an `ErrorCodeReported` issue with the real code (e.g. `0xE0000320 ErrCrypto_BadPassphrase`,
   observed on this machine). With the collector you do **not** need to close OneNote.
   Without the collector, the same detail appears after OneNote next closes (session-end log).
**Undo:** unlock / remove the password.

## 6. Auth expiry — yellow, not red

Delete the token cache (`%LOCALAPPDATA%\OneNoteWatcher\` token file). **Expect:** yellow icon +
"Sign in to Graph" balloon; no red. **Undo:** sign in from the menu.

## 7. OneNote closed — grey, no false alarm

Quit OneNote entirely. **Expect:** grey (idle) icon, no alerts (nothing can sync); the history
file remains readable.

## Verifying the ETW capture itself (diagnostic)

`experiments\etw_probe.ps1` (Admin) captures 60 s of the live providers and greps for sync events —
use it to confirm ETW carries the events on a new machine/build before relying on the collector.

## Notes

- All scenarios target the **Win32 M365** OneNote (`ONENOTE.EXE`), not the UWP "OneNote for
  Windows 10" app.
- Turning Wi-Fi on/off needs no admin; do not use firewall/hosts edits (those would).
- The watcher never interacts with the OneNote UI in any scenario — it only reads files, the event
  log, ETW, and Microsoft Graph.
