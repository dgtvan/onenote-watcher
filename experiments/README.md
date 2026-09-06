# experiments — the probes behind the design

Read-only scripts that produced the measurements in [../docs/detection-design.md](../docs/detection-design.md).
They are kept for two reasons: the claims in the docs stay checkable rather than being assertions,
and `etw_probe.ps1` is the way to validate a **new machine or a new Office build** before trusting the
collector on it.

None of them write to OneNote or change any setting.

| Script | What it answers | Needs |
|---|---|---|
| `etw_probe.ps1` | *Does this machine's OneNote actually stream sync events over ETW?* Captures 60 s of `OfficeLoggingLiblet` while you edit a note, decodes it, and greps for sync events. **Run this first on any new machine.** | Administrator |
| `probe_diagnostic_logs.ps1` | *Which Office diagnostic logs are readable, what sync markers do they contain, and what privileges do I have?* Reproduces the "live log is locked deny-read" finding. | nothing |
| `inspect_search_index.py` | *What does OneNote's local index think the notebook/section/page tree and its timestamps are?* The local half of the cloud-side comparison. | nothing |
| `bond_decoder.py` | *What is in the OTele telemetry store?* Decodes its Bond/1DS payloads. The OTele path is **not** used by the app — this is the tooling for that documented last resort. | nothing |

For day-to-day checking, prefer the built-ins over these scripts:

```powershell
OneNoteWatcher.Collector.exe --audit     # run the real classifier over your diagnostic logs
OneNoteWatcher.exe --graph-check         # compare the local index against Microsoft Graph
```
