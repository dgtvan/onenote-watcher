# Data Sources Reference

Exact formats of every on-disk / OS source the watcher reads. All paths are per-user, none need
elevation. Verified on OneNote 16.0.20326.20112 (Win32, M365), Windows 11 10.0.26200.

## 1. FullTextSearchIndex (local truth)

**Path:** `%LOCALAPPDATA%\Microsoft\OneNote\16.0\FullTextSearchIndex\{<notebook GUID>}{<n>}.db`
**Format:** SQLite 3, rollback-journal mode (`.db-journal` sibling, usually 0 bytes).
**Concurrency:** open read-only with `?mode=ro`; retry on `SQLITE_BUSY` (OneNote holds brief
write locks while indexing). Copying the file is also safe.
**Lifetime:** one file per notebook that has been opened; persists after the notebook is closed
(stale files are harmless — the `Type=4` row's `LastModifiedTime` tells you how old it is).

### `Entities`

| Column | Meaning |
|---|---|
| `Type` | 4 notebook · 3 section group · 2 section · 1 page |
| `GOID` | object id `{guid}{n}` — the `{guid}` prefix is the notebook's index GUID for pages/sections in it |
| `GUID` | object GUID |
| `GOSID` | **stable id, equals the COM/hierarchy `ID` minus the trailing `{B0}`** |
| `ParentGOID` | parent's `GOID` (empty for the notebook root) |
| `GrandparentGOIDs` | ancestor chain |
| `LastModifiedTime` | **FILETIME** (100 ns ticks since 1601-01-01 UTC) |
| `RecentTime` | last-viewed FILETIME (pages only) |
| `Title` | page/section title |
| `Color`, `PinTime`, `ContentRID`, `RootRevGenCount`, `EnterpriseIdentity` | unused here |

```sql
-- newest change per section, with names
SELECT s.Title AS section, s.GOSID,
       MAX(p.LastModifiedTime) AS newest_page, s.LastModifiedTime AS section_lm
FROM Entities s LEFT JOIN Entities p ON p.ParentGOID = s.GOID AND p.Type = 1
WHERE s.Type = 2 GROUP BY s.GOID;
```

FILETIME → UTC: `DateTime.FromFileTimeUtc(value)` in .NET.

**⚠ `LastModifiedTime` is NOT "last content edit" (verified 2026-09-06).** It is re-stamped when a
notebook is (re)opened or re-synced on this PC: every section of *Work* carried
`2026-08-10 09:37:28–41Z` (a 13-second burst) while Graph showed their real edits months earlier,
and the live ETW feed reported the notebook syncing fine. Comparing it absolutely to Graph's
`lastModifiedDateTime` produced 15 false "upload stuck" alerts. Only *movement* of this value while
the watcher runs is meaningful (see `OutcomeDetector`); the ETW sync results are the authority.
The index root `Title` can also be a stale nickname (`Van`) while OneNote/Graph show `Note` — the
resolver prefers the MRU/Graph name.

Other tables (`PageElements`, `EntityContents*`, `Hashtags`, `NoteFlags`) are FTS4 content —
not needed, and contain note text, so never log them.

## 1b. Office MRU cache (notebook OneDrive resource id → name, offline)

`%LOCALAPPDATA%\Microsoft\Office\16.0\MruServiceCache\<identity>\OneNote\Documents_<locale>` — a JSON
array of `{ "FileName": "Work", "ResourceId": "36b934175dc7e3a4!s6f96…", "DocumentUrl": "https://d.docs.live.net/…/Work", … }`.
Turns `Data.NotebookId_ResourceId` into a notebook name without Graph. Note the name is the OneDrive
*folder* name (e.g. `Note`), which can differ from the notebook's display title in the index (`Van`);
the resolver prefers the index title via the rid→GOSID pairs learned from `NotebookSyncResult` events.
Section resource ids do **not** appear anywhere locally (verified) — they resolve through the Graph
section map below.

## 1c. The one root folder (`C:\ProgramData\OneNoteWatcher`)

Executables + `config.ini` + `status.json` + `section-names.json` + `logs\` all live in this single
root (`[etw] shared_dir`, defaulting to the executables' own folder). It is under ProgramData rather
than Program Files because the *unelevated* tray must write there too. `install.ps1 -UpdateOnly`
swaps the binaries in place and keeps your `config.ini`; a full `install.ps1` also replaces
`config.ini` from the repo, saving the old one as `config.ini.bak-<timestamp>`. `status.json`,
`section-names.json` and the logs folder are never touched by either.

| File | Writer | Content |
|---|---|---|
| `status.json` | collector (30 s heartbeat + on change) | `WatcherStatus`: collector alive, OneNote running, internet, last sync activity, per-notebook last activity/result, active issues (full WHERE/WHAT/WHY/FIX) |
| `section-names.json` | tray (after each Graph poll) | `SectionNameMap`: Graph id / id tokens / `section-id={GUID}` / `file:<name>.one` → "Notebook / Section" |
| `processed-sessions.txt` | collector | diagnostic-log session files already back-filled |
| **`logs\`** — *every* log of both processes, one file per day, purged after `log_retention_days` (default 14) | | |
| `logs\sync-history-yyyy-MM-dd.log` | collector | one line per sync result: notebook/section sync results, page uploads/downloads (real-time channel), page sessions, connectivity |
| `logs\collector-yyyy-MM-dd.log` | collector | start/args/profile paths, ETW session start/failures/restarts, backfill, every issue raised/recovered, connectivity changes, 10-min summaries, name-resolution misses |
| `logs\tray-yyyy-MM-dd.log` | tray | start, icon state transitions with reason, every Graph poll (counts, baselined, outcome issues), sign-in/out, OAlerts dialogs, errors |
| `logs\graph-check-<stamp>.txt` | tray (button) | the cloud check report |

Only the DPAPI token cache stays per-user (`%LOCALAPPDATA%\OneNoteWatcher\msal.cache`) — it must not be shared.

### Real-time channel events (what a page edit emits)

A page edit in modern OneNote is uploaded through the **real-time channel**, not a `SectionSyncResult`:

| `EventName` | Keys | Meaning |
|---|---|---|
| `Office.OneNote.Storage.RealTime.NoteItHttpUpload` | `UploadTimeInMs`, `SectionId_ResourceId`, `NotebookId_ResourceId`, `ResourceId_CellId` | a page change reached OneDrive |
| `…RealTime.NoteItHttpDownload` | `TimeToConfirmSyncedWithServerInMs`, same ids | a remote change arrived |
| `…RealTime.NoteItService` | `Error` (`"No error"` when healthy), `OperationWithError`, `DownloadCount`, ids | session summary; non-"No error" = real-time failure (Warn) |

`NotebookSyncResult` / `SectionSyncResult` come from the periodic hierarchy sync (minutes apart). The
watcher counts all of them as "sync activity" so an edit that reached your phone moves the timestamps.

## 2. Microsoft Graph OneNote (server truth)

- Auth: MSAL public client, device-code flow, scope `Notes.Read` (+ `offline_access`), audience
  *Personal Microsoft accounts* (the notebooks live on consumer OneDrive `d.docs.live.net`).
  Token cache encrypted with DPAPI under `%LOCALAPPDATA%\OneNoteWatcher\`.
- `GET https://graph.microsoft.com/v1.0/me/onenote/notebooks?$select=id,displayName,lastModifiedDateTime`
- `GET …/me/onenote/sections?$select=id,displayName,lastModifiedDateTime,parentNotebook`
- `GET …/me/onenote/pages?$select=id,title,createdDateTime,lastModifiedDateTime,parentSection&$top=100`
  (follow `@odata.nextLink`; the pages endpoint is the slow one — poll it less often than sections).
- Timestamps are ISO-8601 UTC.
- Throttling: Graph OneNote is rate-limited per app/user; a 5-minute poll of ~200 pages is far
  below limits. Back off on 429 using `Retry-After`.

Identity mapping: Graph ids are opaque (`0-…!…`), not GOSIDs. Match sections by
`(notebook.displayName, section.displayName)`; pages by `(section, title, createdDateTime≈)`.
Duplicated titles fall back to section-level comparison.

## 3. OTele telemetry store (exact error codes, post-mortem)

**Path:** `%LOCALAPPDATA%\Microsoft\Office\OTele\onenote.exe.db` (+ `-wal`, `-shm`)
**Format:** SQLite 3 in WAL mode.
**Open:** `file:...?mode=ro` works while OneNote runs. **Never open the *copied* db with a SQLite
client if you want to keep its WAL** — closing the last connection checkpoints and deletes it.
Read the `-wal` bytes directly for history.

```sql
CREATE TABLE events (record_id TEXT, tenant_token TEXT NOT NULL, latency INTEGER,
  persistence INTEGER, timestamp INTEGER, retry_count INTEGER DEFAULT 0,
  reserved_until INTEGER DEFAULT 0, payload BLOB);
CREATE INDEX k_latency_timestamp ON events (latency DESC, persistence DESC, timestamp ASC);
CREATE TABLE settings (name TEXT, value TEXT, PRIMARY KEY (name));
```

### Payload: 1DS Common Schema record, Bond Compact Binary v1

Record start marker: `29 03 33 2E 30 49` = field 1 string `"3.0"` then field 2 (name) header.

| Field | Content |
|---|---|
| 1 | `ver` = "3.0" |
| 2 | event name with dots replaced by underscores, e.g. `Office_OneNote_Storage_NotebookSyncResult` |
| 3 | time — **.NET ticks** (100 ns since 0001-01-01 UTC) |
| 5 | iKey `o:<tenant>` |
| 21–33 | ext structs: user(21), device(23: `c:<device id>`), os(24), app(25: `ONENOTE`, build), sdk(32: `EVT-Windows-C++-No-3.7.101.1`) |
| 60 | `"custom"` |
| **70** | `[ { 1: map<string, Value> } ]` — **all `App.*`, `Consent.*`, `Data.*`, `Event.*`, `Session.*`, `User.*` properties** |

`Value` struct: `1` = kind, `3` = string, `4` = int64 (bools: kind 6 + `4`=1), `5` = double,
`6` = GUID bytes. Absent `4` on an int kind means 0.

Field header byte: low 5 bits = type, high 3 bits = id (6 → next byte is id; 7 → next 2 bytes).
Types: 0 STOP, 1 STOP_BASE, 2 BOOL, 3 UINT8, 5 UINT32, 6 UINT64, 8 DOUBLE, 9 STRING, 10 STRUCT,
11 LIST, 13 MAP, 16 INT32, 17 INT64. Varints for lengths/ints, zig-zag for signed.
`experiments/bond_decoder.py` implements exactly this and was validated on 87 records.

### Event catalogue (decoded from this machine)

`Office.OneNote.Storage.NotebookSyncResult` — one per notebook sync attempt

| `Data.` key | Type | Example |
|---|---|---|
| `NotebookErrorCode` | int | `0` (non-zero = the `0xE000…`/`0xE40…` code) |
| `Success` | bool | true |
| `Gosid` | string | `{42AC7CD5-3122-4501-B9DF-254D9AC6D8F6}{1}` — **the notebook GOSID**, joins to the search index |
| `NotebookId_ResourceId` / `NotebookId_FileIdentifier` | string | `36B934175DC7E3A4!s6f96…` (OneDrive resource id) |
| `SyncDestinationType` | string | `OneDrive` / `Local` |
| `LastSuccessfulSync`, `LastAttemptedSync`, `LastBackgroundSync`, `LastNotebookViewedDate` | .NET ticks | |
| `TimeSinceLastSuccessfulSync`, `TimeSinceLastAttemptedSync`, `ExecutionTime` | ms | |
| `IsNotebookErrorSuppressed`, `IsSectionErrorSuppressed`, `IsSectionErrorUnexpected`, `IsCachedErrorSuppressed`, `IsCachedErrorUnexpected` | bool | |
| `IsBackgroundSync`, `SyncWasUserInitiated`, `SyncWasFirstInSession`, `IsUsingRealtimeSync`, `InitialReplicationInSession` | bool | |
| `NeedToRestartBecauseOfInconsistencies`, `ReplicatingAgainBecauseOfInconsistencies` | bool | |
| `SyncId`, `IdentityType`, `NotebookType`, `TenantId` | int/string | |

`Office.OneNote.Storage.SyncScore` — OneNote's **background replication scan** noting what it skipped.
**Not a sync failure and never an issue** (verified 2026-09-06: `ErrCrypto_BadPassphrase` fires at every
session start / periodic scan for a locked password-protected section that nobody touched; it carries no
notebook/section id and has no clearing event). Recorded in history as `SCORE …` only.

| `Data.` key | Example |
|---|---|
| `Error_Code` | `3758097184` = `0xE0000320` |
| `Error_Description` | `ErrCrypto_BadPassphrase` |
| `Error_Tag` | `bbv2e` |
| `Error_Type` | `Win32Error` |
| `Source` | `Storage.Replication.Fishbowl` |
| `fishbowlType`, `idsFishbowl` | `SectionFishbowl`, int |

`Office.OneNote.Storage.PageSyncSession` — per active page: `ActivePageGOID`,
`ErrorState_TimeInActiveSyncErrorState`, `ErrorState_TimeInHierarchySyncErrorState`,
`PageSyncState_TimeConnected/NotConnected/Uploading/RequestingDownload`, `TimeUpToDate`, `TimeTotal` (ms).

`Office.System.SystemHealthErrorsWithTag` — `ErrorGroup`, `Trackback`, `Count`, `FirstTimeStamp`, `EndTime`.

Names seen but not decoded (no payload captured yet): `Office.OneNote.BackupTriggeredByStuckContentError`,
`Office.OneNote.ReadOnlyTriggeredByErrorDuration`, `Office.OneNote.Indexer.SqlIndexCorrectingInconsistency`.

**Behaviour:** rows exist only while upload is pending. Measured: 0 rows over 115 s including two
forced syncs; WAL last written at app exit. Treat as an at-startup / on-exit source.

## 4. `OAlerts` Event Log (dialog errors, real-time)

- Log name `OAlerts`, provider `Microsoft Office 16 Alerts`, Event ID 300, Level Information.
- Message: `Microsoft OneNote <dialog text> P1: <numeric id> P2: <build> P3: <tag> P4:`
- Sample (2026-08-10): `Microsoft OneNote We couldn't open that location. It might not exist or
  you might not have permission to open it. Please contact the owner of
  onenote:https://d.docs.live.net/36b934175dc7e3a4/Van @ Dgtvan_bk_6August2k26`
- Subscribe: `new EventLogWatcher(new EventLogQuery("OAlerts", PathType.LogName,
  "*[System[Provider[@Name='Microsoft Office 16 Alerts']]]"))`, filter on message prefix.
- Also logs non-error dialogs ("Are you sure you want to move this section…") — classify by text.

## 5. Office diagnostic logs (exact error codes, post-session — preferred over OTele)

- `%LOCALAPPDATA%\Temp\Diagnostics\ONENOTE\Primary<seq>_<sessionGUID>.log` (two files per
  session share the GUID) and `Additional\Additional<seq>_<sessionGUID>.log` (empty so far).
- 16 MiB preallocated, NUL-padded → **strip `\0`** before parsing. Real content: 10 KB for a
  short session, ~650 KB for a working day.
- The **active session's files are locked deny-read**; a session's log becomes readable only
  after OneNote exits. Identify sessions by the GUID in the file name; remember which GUIDs have
  been processed.
- Row format (TSV): `Timestamp\tProcess\tTID\tArea\tCategory\tEventID\tLevel\tMessage\tCorrelation\r`
  e.g. `09/04/2026 15:54:35.932\tONENOTE (0x10A8)\t0xC624\tMicrosoft OneNote\tTelemetry Event\tb7vzq\tMedium\tSendEvent {…}\t`
  The `Timestamp` column is **local time**; use the JSON `"Time"` (UTC ISO-8601) instead.
- `Message` = `SendEvent ` + a flat JSON object: `EventName`, `Flags`, `InternalSequenceNumber`,
  `Time`, then `Data.*` keys (already flattened — no Bond decoding needed).

Sync events and their `Data.*` keys (all observed on this machine):

| `EventName` | Keys |
|---|---|
| `Office.OneNote.Storage.NotebookSyncResult` | `NotebookErrorCode` (int), `Success`, `Gosid`, `NotebookId_ResourceId`, `NotebookId_FileIdentifier`, `NotebookId_Provider`, `SyncDestinationType`, `LastSuccessfulSync`/`LastAttemptedSync`/`LastBackgroundSync`/`LastNotebookViewedDate` (ISO strings), `TimeSinceLastSuccessfulSync`/`TimeSinceLastAttemptedSync`/`ExecutionTime` (ms), `IsNotebookErrorSuppressed`, `IsSectionErrorSuppressed`, `IsSectionErrorUnexpected`, `IsCachedErrorSuppressed`, `IsCachedErrorUnexpected`, `IsBackgroundSync`, `SyncWasUserInitiated`, `SyncWasFirstInSession`, `IsUsingRealtimeSync`, `InitialReplicationInSession`, `NeedToRestartBecauseOfInconsistencies`, `ReplicatingAgainBecauseOfInconsistencies`, `SyncId`, `NotebookType`, `TenantId`, `IdentityType` |
| `Office.OneNote.Storage.SectionSyncResult` | `Error_Code` (uint32, e.g. `3758096477` = `0xE000005D`), `Error_Description` (`ErrFilePendingRename`), `Error_Tag`, `Error_Type` (`Win32Error`), `ErrorLast`, `IsErrorTransient`, `IsErrorSuppressed`, `IsErrorUnexpected`, `Success`, `SectionResourceId_ResourceId` (OneDrive item id — joins to Graph), `SectionResourceId_FileIdentifier`, `SectionPath` / `NotebookPath` (hashed), `UnmappedGosid`, `IsEncrypted`, `NotebookId_ResourceId`, `SyncDestinationType`, `SyncId`, `IsBackgroundSync` |
| `Office.OneNote.Storage.SyncScore` | `Error_Code`, `Error_Description`, `Error_Tag`, `Error_Type`, `Source` (`Storage.Replication.Fishbowl`), `fishbowlType`, `idsFishbowl`, `IsUsingRealtimeHierarchySync` |
| `Office.OneNote.Storage.PageSyncSession` | as in §3 |
| `Office.OneNote.Storage.ConnectivityChanged` | `InternetConnectivityNowAvailable` |

Error fields are absent when there is no error (`Error_Code` key missing ⇒ success). Display codes
as `0x{Error_Code:X8}`; `0xE000xxxx` are OneNote store errors, `0xE40xxxxx` service errors (see the
Microsoft "Fix issues when you can't sync OneNote" article for the public mapping).

## 5b. Section and notebook identifiers across event kinds (verified)

Decoded from real payloads on 2026-09-06. This matters because errors are matched to the successes
that clear them by these identifiers — if two event kinds spelled the same section differently, an
error could never be cleared.

| Event | Section field | Notebook field |
|---|---|---|
| `Storage.SectionSyncResult` | `Data.SectionResourceId_ResourceId` = `36B934175DC7E3A4!s8d49…` | `Data.NotebookId_ResourceId` = `36B934175DC7E3A4!626` |
| `Storage.RealTime.NoteItService` | `Data.SectionId_ResourceId` = `36B934175DC7E3A4!s7720…` | `Data.NotebookId_ResourceId` = `36B934175DC7E3A4!626` |
| `Storage.NotebookSyncResult` | — | `Data.NotebookId_ResourceId`, plus `Data.Gosid` = `{GUID}{1}` |

**Confirmed: section sync and the real-time channel use the identical `!s<hex32>` resource-id
spelling**, so `section:` and `section-rt:` scope keys line up between them and coverage works. Do not
"fix" this by normalising ids without re-checking — the field *names* differ
(`SectionResourceId_ResourceId` vs `SectionId_ResourceId`) even though the values match.

Other spellings seen for the same section, all present as keys in `section-names.json`: the bare
`<hex32>` token, and the `0-`/`0|` prefixed `FileIdentifier` form. `SectionNameMap` indexes every one.

`NotebookSyncResult` also carries `IsSectionErrorSuppressed` / `IsSectionErrorUnexpected`, which is the
evidence that a notebook sync can report success while its sections failed — see `docs/fail-closed.md`.

## 6. Process presence

`ONENOTE.EXE` (`C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE`). Presence gates the
alert rules (nothing can sync while it is closed) and the OTele post-mortem read (run it when the
process *disappears*). Do **not** create `OneNote.Application` COM objects to check — that starts
the app.

## Epochs cheat-sheet

| Where | Epoch | Convert |
|---|---|---|
| FullTextSearchIndex `LastModifiedTime`, `RecentTime` | 1601-01-01 (FILETIME) | `DateTime.FromFileTimeUtc` |
| OTele record field 3, `Data.Last*Sync` | 0001-01-01 (.NET ticks) | `new DateTime(ticks, DateTimeKind.Utc)` |
| Graph `lastModifiedDateTime` | ISO-8601 | `DateTimeOffset.Parse` |
| OneNote error codes | — | `Error_Code` is an unsigned 32-bit HRESULT-style value; print as `0x{0:X8}` |
