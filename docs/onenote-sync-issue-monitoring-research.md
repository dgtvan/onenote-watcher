# Reading OneNote's Internal Sync Status Programmatically: A Technical Assessment

## TL;DR
- **The single most practical route is not injection or COM at all — it is tailing OneNote's own plain-text Office diagnostic logs under `%LOCALAPPDATA%\Temp\Diagnostics\` (OneNote subfolder), which contain the exact sync error codes (e.g. `"Error.Code":"0xE000012E"`, `"Error.Type":"Win32Error"`) as readable JSON `SendEvent` records.** This captures the silent-failure case directly, survives Office auto-updates (it's a stable Office-wide logging schema), and needs no injection.
- **UI Automation of the "Shared Notebook Synchronization" dialog is the best fully-supported, non-invasive fallback** for reading notebook names, sync state, and the visible error code/message strings — but it requires the dialog to be open and OneNote's Office ribbon/dialogs expose UIA text imperfectly, so it is fragile for silent polling.
- **DLL injection + API hooking is technically possible but is the worst option**: no public OneNote symbols (Office PDBs are not on the Microsoft symbol server), offsets break every monthly Office update, and Office anti-tampering/Code Integrity Guard actively fights injection. The documented COM `IApplication` interface confirmed does NOT expose sync error codes.

## Key Findings

**The documented COM API is a confirmed dead end for error codes.** OneNote's `Application.GetHierarchy` returns an XML hierarchy whose node attributes are limited to things like `name`, `ID`, `path`, `lastModifiedTime`, `color`, `isCurrentlyViewed`, `isUnread`, and `nickname`. There is no sync-status or sync-error-code attribute anywhere in the documented object model, and `SyncHierarchy` merely *triggers* a sync — it does not report per-notebook error state. This is confirmed by the Microsoft Learn "Application interface (OneNote)" reference and by the OneNote XSD schema attributes that appear across all public examples.

**OneNote emits sync errors to disk as readable telemetry.** When a notebook fails to sync, OneNote writes a structured record such as `"Error.Code": "0xe000012e", "Error.Type": "Win32Error", "Error.Description": "jerrcFileNodeFileCorrupt_FileNodeListChunkMalformedRanOffEnd"`. This is the Message field of an Office diagnostic-log row. These logs live under `%LOCALAPPDATA%\Temp\Diagnostics\` (per-app subfolders), with the copy uploaded to Microsoft in `%temp%\Diagnostics\UploadCache`. Microsoft documents them as plain text viewable in a text editor, with a tab-structured schema (Timestamp, Process, TID, Area, Category, EventID, Level, Message, Correlation) where the Message column carries a JSON `SendEvent {...}` payload.

**The error codes form a stable, documented namespace.** OneNote sync error codes are `0xE000xxxx` / `0xE40xxxxx` HRESULT-style values (e.g. `0xE000002E` Out of Sync with Store, `0xE000005E` ReferencedRevisionNotFound, `0xE00015E0` insufficient space, `0xE0000796` Quota Exceeded, `0xE000078B` Server Timeout, `0xE4010641`/`0xE4010690` can't reach service, `0xE40105F9` unsupported client version). Microsoft Support's article "Fix issues when you can't sync OneNote" is the matching-code reference; it instructs the user to "Select File > Info > View Sync Status. In the Shared Notebook Synchronization dialog box that opens, check for any error code and message displayed for your notebook, and then check the list of error codes elsewhere in this article for a matching result." Internally each code also has a `jerrc*` symbolic description string.

## Details

### 1. UI Automation / MSAA — VERDICT: Best supported non-invasive route, but fragile for silent polling

Microsoft UI Automation (UIA) is the same platform screen readers use, and OneNote *is* screen-reader-instrumented: Microsoft's own "Use a screen reader to explore and navigate OneNote" documents that the ribbon row exposes a "Sync status" button, and the navigation pane exposes the notebook picker and section/page lists as navigable elements. That means the sync glyph and the "Shared Notebook Synchronization" dialog contents are, in principle, reachable through the UIA tree.

- **Tooling:** Use `inspect.exe` or Accessibility Insights for Windows to enumerate the actual tree first, then `FlaUI` (FlaUI.UIA3) or `System.Windows.Automation` in C# to read it. Search elements by `AutomationId` (locale-independent) where present; fall back to `Name`/ControlType.
- **Reality check / gotchas:** Office dialogs are Win32/legacy, not WPF, so many controls lack a stable `AutomationId` — you often must match on displayed text, which is locale-dependent and changes with UI updates. The "Shared Notebook Synchronization" dialog must be *open* (`File > Info > View Sync Status`) for its per-notebook rows and error strings to exist in the tree; UIA cannot read a dialog that isn't rendered. The persistent per-notebook sync glyph in the nav pane (green/red/warning triangle) may be drawn (GDI/DirectUI) rather than exposed as a discrete named element with a readable state — this must be verified with inspect.exe on the target build.
- **Feasibility:** Feasible and fully supported (no anti-tampering issues). To poll silently you would have to programmatically open the sync-status dialog on a timer, which is intrusive to the user. Survives Office updates reasonably well structurally, but text/automation-ID matching is version- and locale-fragile.

### 2. DLL injection + API hooking — VERDICT: Exotic, most fragile, effectively a dead end for a production monitor

- **Frameworks:** Microsoft Detours, MinHook, EasyHook, or Frida can all inject into and hook a running x64 process in principle. Ironically, Office itself ships Detours internally (the App-V subsystem `AppvIsvSubsystems64.dll` is Detours-injected into Office processes, per Deep Instinct's reverse-engineering writeup), proving injection is mechanically possible.
- **Blocking difficulties (all confirmed concerns):**
  - **No symbols.** Microsoft's public symbol server (`msdl.microsoft.com/download/symbols`) hosts Windows PDBs but **not Office PDBs**. This is long-standing: Microsoft support staff state "symbols for Office are not made available to customers … never were." So you'd be reverse-engineering `onenote.exe`/`mso*.dll` blind with IDA/Ghidra to find the sync-status function.
  - **Volatile offsets.** Microsoft 365 Office updates monthly (build numbers like 16.0.15726.xxxxx roll constantly). Any hard-coded function offset or vtable index breaks on nearly every update, plus ASLR.
  - **Anti-tampering.** Office can run with Code Integrity Guard (blocks loading non-Microsoft-signed images), forcing manual/reflective mapping to inject at all; and injecting into Office is a red flag for Defender/EDR (MITRE ATT&CK catalogs Office process injection as an attack technique).
  - **x64 calling conventions** and undocumented internal C++ interfaces make safe hooking hard.
- **Modules that plausibly handle sync:** `onenote.exe` itself plus `mso*.dll`/`mso20win32client.dll`/`mso30win32client.dll` and OneNote's own sync/OneStore code; the actual sync/FSSHTTP logic is not in a cleanly-named separate DLL you can proxy. There is no documented internal interface. `msosync.exe`/Upload Center (below) is a separate mechanism that does **not** cover OneNote.
- **DLL proxying / decompile-and-proxy:** Same symbol and update-cadence problems, plus you'd need to identify a specific DLL whose export you can meaningfully intercept — none is documented for sync status. Not viable.

### 3. COM / undocumented interfaces — VERDICT: Documented API confirmed insufficient; no public undocumented sync interface found

Beyond `IApplication`, OneNote does expose COM events, but the documented events concern hierarchy/navigation changes, not sync error state. The `OneNote.Application` COM object (`New-Object -ComObject OneNote.Application`) and its typelib do not surface sync error codes. Note an important distinction: the Office-wide `Office.Sync` object (`MsoSyncStatusType`, `msoSyncStatusError`, `ErrorType`) exists in Word/Excel/PowerPoint's object model for document-workspace sync — but OneNote is not a `Document`/`Workbook`/`Presentation` host and does not expose that `Sync` object for its notebook sync. No registry-registered undocumented OneNote sync CLSID/IID surfaced in research. Treat any undocumented COM sync interface as unconfirmed/nonexistent for practical purposes.

### 4. Local cache / filesystem signals — VERDICT: Useful corroborating signal, not a clean error source

- **Location:** `%LOCALAPPDATA%\Microsoft\OneNote\16.0\` for the 2016/M365-lineage Win32 app. Contains a `cache\` subfolder, `OneNoteOfflineCache.onecache` (and historically GUID-named `.bin`/`{GUID}` files under `OneNoteOfflineCache_Files`), plus `16.0\Backup\`.
- **Format:** Proprietary binary. The public reverse-engineering that exists is for the *notebook* file format (`.one`/`.onetoc2`) — see `msiemens/onenote.rs`, which implements `[MS-ONESTORE]` (OneNote Revision Store), `[MS-FSSHTTPB]` (binary packaging), and `[MS-ONE]` (file format). These specs describe *content*, not a documented "pending upload queue" or "sync error" field.
- **What you can infer:** A growing/locked `.onecache`, files that won't optimize, or a persistently large pending-changes cache correlate with a stuck sync. `FileSystemWatcher` on the cache can reveal that *changes are being written locally* but you cannot cleanly read "N items pending upload" or an error code from the binary. There is a documented registry breadcrumb: OneNote writes `LastCacheOptimizeSuccessTime` under the OneNote\General registry branch, per a Microsoft archived engineering blog — a weak, indirect health signal only.
- Feasible as a secondary heuristic; noisy; does not cleanly capture the silent-error case on its own.

### 5. ETW / diagnostics — VERDICT: The on-disk diagnostic log (5a) is the winner; real-time ETW (5b) is experimental

**5a. On-disk Office diagnostic logs — strongest confirmed signal.** As above: `%LOCALAPPDATA%\Temp\Diagnostics\<APP>\`, plain text, JSON `SendEvent` records containing the `0xE000xxxx` sync error codes and their `jerrc*` descriptions. Microsoft documents the format ("Overview of diagnostic log files for Office") and confirms OneNote is in scope. Approach: `FileSystemWatcher` + tail-and-parse the newest file, match Message-column JSON for `"Error.Type":"Win32Error"` / OneNote sync EventNames. **Caveats:** (a) generation is gated on Office diagnostic-data / connected-experiences settings being enabled — a monitor cannot assume the files always exist; (b) there may be write-buffering latency between the sync failure and the line hitting disk (verify empirically); (c) files rotate (24h retention, size caps).

**5b. Real-time ETW.** Office (incl. OneNote) emits TraceLogging/ETW events that feed the DiagTrack "Diagtrack-Listener" session. The BSI/ERNW "SiSyPHuS Win10 – Work Package 4: Telemetry" study (Version 1.0) confirms this, stating that "An Office application may produce diagnostic events using ETW providers that are associated with Diagtrack-Listener," which are then consumed by the Connected User Experiences and Telemetry (DiagTrack) service. A C#/.NET consumer can subscribe in real time using `Microsoft.Diagnostics.Tracing.TraceEvent` (the PerfView library), but requires elevation or Performance Log Users membership. **Critical gap:** there is **no publicly documented OneNote/Office sync ETW provider GUID**; the name "Microsoft-Office-OneNote" is plausible but unverified, and Office provider GUIDs are dynamic (mapped via Microsoft-controlled `utc.app.json`). You would have to discover the live provider on the target machine (`logman query providers | findstr /i office`) and empirically confirm it carries the sync events. Treat as experimental.

**5c. Do NOT use `%temp%\OneNote.log`.** Setting `HKCU\Software\Microsoft\Office\16.0\OneNote\Options\Logging` `EnableLogging=1` (with `ttidLogObjectModel`/`ttidLogObjectModelAddins`) writes `%temp%\OneNote.log`, but Microsoft's own archived blog confirms this logs **COM/object-model API calls only**, not sync errors. Wrong file for this purpose.

**5d. OneNote Client Diagnostics Tool** (`microsoft.com/download` id=53336; installs to `C:\Program Files (x86)\OneNoteDiagnostics\`) is a *collector/uploader* that packages the existing on-disk logs into a `.cab` — it confirms the logs exist and contain sync/error data but is not itself a real-time API.

### 6. Registry — VERDICT: Weak; no reliable live sync-error value

`HKCU\Software\Microsoft\Office\16.0\OneNote\` holds `OpenNotebooks` (notebook locations) and General-branch values like `LastCacheOptimizeSuccessTime`. Research surfaced no registry value that reliably records last-sync-time or a live per-notebook sync error code. Registry is useful for enumerating which notebooks exist/are cloud-backed, not for detecting failure. Do not rely on it as the primary signal.

### 7. Network-level — VERDICT: Possible but heavy and indirect

OneNote syncs to OneDrive/SharePoint via **FSSHTTP/Cobalt** (`[MS-FSSHTTP]`/`[MS-FSSHTTPB]`) — incremental binary deltas over SOAP/HTTP — with consumer notebooks hitting `d.docs.live.net`/`onenote.com` endpoints. A local proxy (with TLS interception) or an ETW network trace could in theory observe repeated non-200/error responses on these calls. This is invasive (certificate trust, privacy), complex to correlate to a specific notebook, and does not give you OneNote's own error code. Suitable only as a last resort or for a network-monitoring context, not an app-level health check.

### 8. Existing projects & Upload Center relevance

- **Microsoft Office Upload Center (`msosync.exe`)** manages pending uploads for Word/Excel/PowerPoint Office Document Cache files — it does **NOT** cover OneNote, which uses its own sync engine (a TechNet thread confirms OneNote sync is independent of MSOSYNC). So Upload Center logs are irrelevant here.
- **`msiemens/onenote.rs`** (Rust) and the `onenote_parser` crate — best public reverse-engineering of the OneNote/OneStore/FSSHTTPB *file* formats. Useful if you go the cache-parsing route; does not parse sync state.
- **OneNote COM helper projects** (e.g. `matthew-zhang-306/OneNote-x-CSharp`, `richie5um/OneNoteJournal`, various PowerShell `GetHierarchy` samples) show the standard COM interop but none expose sync errors — corroborating that the API can't do this.
- No mature open-source "OneNote sync health monitor" that reads OneNote's own error state was found; existing "monitors" are OneDrive-file-level and explicitly break on `.one` files (see `abraunegg/onedrive` issue #2578, where `--display-sync-status` fails on OneNote sections).

## Recommendations

**Stage 1 — Build on the diagnostic log (primary, do this first).**
Implement a `FileSystemWatcher` + tail-parser over `%LOCALAPPDATA%\Temp\Diagnostics\` (filter to the OneNote process subfolder). Parse each row's Message JSON; raise your loud alert whenever you see a OneNote `SendEvent` with `"Error.Type":"Win32Error"` and an `"Error.Code":"0xE000…"` / `"0xE40…"`. Maintain a small allow/deny map of codes:
- Fire immediately on hard-fail codes — e.g. `0xE000005E` (ReferencedRevisionNotFound — "a section of one or more notebooks fails to sync," fixed by a manual Sync Now) and `0xE00009C8` (a sync interruption the OneNote team publicly acknowledged in late 2024, remedied by triggering a manual sync after their fix).
- Treat transient codes like `0xE4020040` ("store busy, retry later") with a retry-count/time threshold before alerting.

This is the only approach that captures the *silent* failure with the actual error code and message, without user interaction.
- **Precondition to verify on target:** confirm Office diagnostic data / connected experiences are enabled so the logs generate; if an enterprise policy disables them, fall back to Stage 2.

**Stage 2 — Add UIA as a corroborating/fallback probe.**
Use FlaUI (UIA3) to (a) read the nav-pane sync glyph state if it proves exposed, and (b) on demand, open `File > Info > View Sync Status` and scrape per-notebook rows/error strings. First run `inspect.exe`/Accessibility Insights against your exact OneNote build to record the real AutomationIds/ControlTypes; guard all matches with try/catch and version checks. Use this to confirm a Stage-1 alert and to get the human-readable notebook name.

**Stage 3 — Cache/registry heuristics as a liveness backstop.**
Watch `%LOCALAPPDATA%\Microsoft\OneNote\16.0\cache` for a pending-changes backlog that isn't draining, and read `LastCacheOptimizeSuccessTime`. Use only to detect "sync appears stalled" when Stages 1–2 are unavailable.

**Do NOT invest in:** DLL injection/hooking or DLL proxying: no symbols, offsets break monthly, anti-tampering/EDR friction, high maintenance — reserve only for a research spike, never production. Skip `%temp%\OneNote.log` (wrong data). Skip Upload Center (doesn't cover OneNote). Treat real-time ETW as an optional enhancement only after you discover and verify a provider GUID on your target.

**Thresholds that change the plan:**
- If diagnostic logging is policy-disabled across your fleet → UIA (Stage 2) becomes primary.
- If you need sub-second latency **and** confirm a working Office ETW provider GUID carries the sync events → promote ETW (5b) above log-tailing.
- If **OneNote for Windows 10** (the UWP app) rather than the Win32 app is in scope → most of this changes (different cache path `%LOCALAPPDATA%\Packages\Microsoft.Office.OneNote_8wekyb3d8bbwe`, no classic COM). Per Microsoft's M365 Message Center announcement (updated March 20, 2025), OneNote for Windows 10 reaches end of support on October 14, 2025, after which "the app will enter read-only mode … editing, creating, or syncing content will no longer be supported" — so target the Win32 `ONENOTE.EXE` regardless.

## Caveats
- **Version/build dependence:** Everything here is tied to the Win32 M365/2016-lineage `ONENOTE.EXE`. Paths (`16.0`), UIA trees, and internal structures differ for OneNote for Windows 10/UWP and across Office channels.
- **Diagnostic-log availability is not guaranteed** — it depends on Office diagnostic-data settings and may be reduced by enterprise privacy policy; write latency and rotation are not fully documented and must be validated empirically.
- **No confirmed OneNote ETW provider GUID and no confirmed undocumented COM sync interface** were found in public sources; those paths are unverified.
- **Injection/hooking** is not only fragile but may violate Office licensing/anti-tampering expectations and will likely trigger security tooling.
- Some sync error-code meanings come from community/support pages rather than a single authoritative enumeration; validate specific codes against the current Microsoft Support "Fix issues when you can't sync OneNote" article.