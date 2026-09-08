using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using OneNoteWatcher.Core.Model;

namespace OneNoteWatcher.Core.Parsing;

/// <summary>
/// Parses and classifies OneNote telemetry from the <c>SendEvent {…}</c> JSON emitted both live over
/// ETW and into the Office diagnostic log.
///
/// FAIL-CLOSED CONTRACT (docs/fail-closed.md): a sync is treated as successful ONLY on an
/// explicit positive signal. Every <c>Office.OneNote.Storage.*</c> event is parsed — including ones
/// this build has never seen — and anything that cannot be proven successful is surfaced rather than
/// dropped. Non-Storage <c>Office.OneNote.*</c> events are picked up when they carry a failure signal
/// in their name or fields, so a future "…Failed"/"…Blocker"/"…Error" event is caught without a code change.
/// </summary>
public static class SyncEventJson
{
    public const string SendEventPrefix = "SendEvent ";
    private const string OneNoteMarker = "\"Office.OneNote.";
    private const string StoragePrefix = "Office.OneNote.Storage.";

    /// <summary>Cap on the payload kept for an unclassified event — enough to classify it, bounded.</summary>
    private const int MaxRawPayload = 2000;

    /// <summary>Words in an event NAME that mean "something went wrong" even with no error field.</summary>
    private static readonly Regex FailureWordsInName = new(
        "Fail|Error|Blocker|Blocked|Stuck|Corrupt|ReadOnly|Conflict|Denied|Unauthori|Expired|Rejected|Abort|Crash|Inconsistenc",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Names where a failure word does NOT describe a failure. A "request blocker" is OneNote's
    /// internal blocking-wait handle — instantiating one means a caller is waiting for a download,
    /// not that sync is blocked. Observed 2026-09-06: FdoDownloadRequestBlockerInstantiated fired
    /// between a healthy PAGE-SESSION and a healthy REALTIME on the same section, in the same second.
    /// Deliberately narrow: a plain "…SyncBlockerInstantiated" is still treated as a failure signal.
    /// </summary>
    private static readonly Regex FailureWordExceptions = new("RequestBlocker", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>Field names that carry an explicit success flag.</summary>
    private static readonly string[] SuccessFields = ["Data.Success", "Data.IsSuccess", "Data.Succeeded", "Data.WasSuccessful"];

    /// <summary>Field-name pattern for numeric error codes (Error_Code, NotebookErrorCode, SH_ErrorCode, …).</summary>
    private static readonly Regex ErrorCodeField = new(@"^Data\.[A-Za-z0-9_.]*Error_?Code$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>An 8-digit hex code embedded in free-text error output, e.g. "(0xE000002E)".</summary>
    private static readonly Regex EmbeddedHexCode = new(@"0x([0-9A-Fa-f]{8})(?![0-9A-Fa-f])", RegexOptions.Compiled);

    /// <summary>
    /// Field-name pattern for the HTTP status of a service call. Only names that say "Http" outright —
    /// a bare "Status"/"StatusCode" is used for too many non-HTTP things to read as an outcome.
    /// </summary>
    private static readonly Regex HttpStatusField = new(@"^Data\.Http(Status|StatusCode|ResponseCode)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>Field-name pattern for error text.</summary>
    private static readonly Regex ErrorTextField = new(@"^Data\.(Error|Error_Description|Error_Type|OperationWithError|FailureReason|ErrorMessage)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>Values in an error-text field that actually mean "fine".</summary>
    private static readonly string[] NonErrorTexts = ["", "no error", "none", "0", "success", "ok", "noerror"];

    /// <summary>
    /// Known non-outcome telemetry: recorded in history, never alerted. Each entry is a deliberate,
    /// evidence-backed decision — not a silent drop.
    /// </summary>
    private static readonly Dictionary<string, string> DiagnosticEvents = new(StringComparer.Ordinal)
    {
        // OneNote's background replication scan. Fires at every session start / periodic pass listing what
        // it SKIPPED (e.g. a locked password-protected section). No notebook/section id, no clearing event
        // → alerting on it is a permanent false alarm. Verified 2026-09-06.
        ["Office.OneNote.Storage.SyncScore"] = "background replication scan (reports skipped items, not a sync attempt)",
        // Per-page timing metric; its ErrorState_Time* fields are handled explicitly below.
        ["Office.OneNote.Storage.PageSyncSession"] = "per-page timing metric",
        // Connectivity transition; handled as its own signal, not an outcome.
        ["Office.OneNote.Storage.ConnectivityChanged"] = "connectivity transition",
    };

    /// <summary>
    /// Known telemetry that reports no outcome of its own — benign UNLESS the payload itself carries
    /// error evidence.
    ///
    /// This is NOT <see cref="DiagnosticEvents"/>, and the difference is the whole point: that table is
    /// consulted BEFORE the error checks (SyncScore reports errors it skipped, so its codes must not be
    /// read as a failed sync), which means anything listed there has its error fields swallowed. These
    /// events have no such exemption — the lookup happens AFTER the error checks, so a genuinely failed
    /// attachment download still surfaces as a failure with its code.
    ///
    /// "FDO" = File Data Object: an embedded or attached file (inserted file, image, printout, recording)
    /// stored apart from the page XML and fetched over the same Cobalt/FSSHTTP protocol as page content.
    /// All five were observed 2026-09-06 23:01, each one bracketed by a healthy PAGE-SESSION and a healthy
    /// REALTIME on the same section within the same second, and none carried an error code, error text or
    /// a success flag — see docs/fail-closed.md.
    /// </summary>
    private static readonly Dictionary<string, string> BenignUnlessError = new(StringComparer.Ordinal)
    {
        ["Office.OneNote.Storage.RealTime.FileDataObjectDownload"] = "an embedded-file (attachment) download, reported without an outcome",
        ["Office.OneNote.Storage.RealTime.DownloadFdoViaCobalt"] = "an embedded-file download over Cobalt/FSSHTTP, reported without an outcome",
        ["Office.OneNote.Storage.RealTime.DownloadFdoStats"] = "timing/size statistics for an embedded-file download",
        ["Office.OneNote.Storage.RealTime.FdoDownloadRequestBlockerInstantiated"] = "an internal blocking-wait handle for an embedded-file download was created — not a sync blocker",
        ["Office.OneNote.Storage.RealTime.DoesNotebookSatisfyNoteItPrerequisites"] = "a capability check asking whether the notebook can use the real-time channel",
        // OneNote's startup sync gate. Evidence (2026-09-06 08:25, 2026-09-06 11:58, 2026-09-08 10:20):
        // it fires exactly ONCE per OneNote launch, one second after the connectivity ONLINE events, and
        // is followed within 0-2 s by a successful PAGE-DOWNLOAD and then a clean session. Its payload —
        // captured by the UNCLASSIFIED PAYLOAD log — carries NO Data.* fields whatsoever: no notebook, no
        // section, no code, no reason. An event that cannot say what is blocked, contradicted every time
        // by content moving seconds later, is the sync engine being constructed, not a sync failure.
        // A real blockage still surfaces: it would come through the Storage events that carry identity.
        ["Office.OneNote.Storage.RealTime.SyncBlockerInstantiated"] = "OneNote's startup sync gate being constructed (no payload, always followed by successful sync)",
    };

    /// <summary>Cheap ETW hot-path pre-filter: any OneNote SendEvent. (≈640 Office events/s, few are OneNote.)</summary>
    public static bool LooksLikeSyncSendEvent(string? message) =>
        message is not null
        && message.StartsWith(SendEventPrefix, StringComparison.Ordinal)
        && message.Contains(OneNoteMarker, StringComparison.Ordinal);

    /// <summary>
    /// Parse and classify. Returns null only for events that are provably out of scope (a non-Storage
    /// OneNote event with no failure signal) or non-OneNote payloads. A malformed Storage payload is
    /// returned as <see cref="SyncOutcome.Unknown"/> rather than dropped — see <paramref name="malformed"/>.
    /// </summary>
    public static SyncEvent? TryParse(string? message) => TryParse(message, out _);

    /// <param name="malformed">Set when the message looked like a OneNote SendEvent but could not be parsed.</param>
    public static SyncEvent? TryParse(string? message, out bool malformed)
    {
        malformed = false;
        if (string.IsNullOrEmpty(message)) return null;
        var json = message.StartsWith(SendEventPrefix, StringComparison.Ordinal) ? message[SendEventPrefix.Length..] : message;
        json = json.Trim().TrimEnd('\r', '\n', '\t');
        if (json.Length == 0 || json[0] != '{') return null;

        JsonDocument doc;
        try
        {
            // The diagnostic log's Message column is followed by a tab and the Correlation id, so the
            // payload is "{…}\t<guid>". Read exactly one JSON value and ignore anything after it.
            var reader = new Utf8JsonReader(System.Text.Encoding.UTF8.GetBytes(json),
                new JsonReaderOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip });
            if (!JsonDocument.TryParseValue(ref reader, out doc!)) throw new JsonException("no JSON value");
        }
        catch (JsonException)
        {
            // fail-closed: if it named a OneNote Storage event we could not read, say so loudly
            if (json.Contains(StoragePrefix, StringComparison.Ordinal))
            {
                malformed = true;
                return new SyncEvent
                {
                    EventName = ExtractName(json) ?? "Office.OneNote.Storage.<unparseable>",
                    Kind = SyncEventKind.UnrecognisedStorage,
                    Outcome = SyncOutcome.Unknown,
                    Time = DateTimeOffset.UtcNow,
                    ClassificationReason = "the payload could not be parsed, so success cannot be confirmed",
                    RawPayload = json.Length <= MaxRawPayload ? json : json[..MaxRawPayload] + "…(truncated)",
                };
            }
            return null;
        }

        using (doc)
        {
            var root = doc.RootElement;
            if (!root.TryGetProperty("EventName", out var nameEl)) return null;
            var name = nameEl.GetString();
            if (name is null || !name.StartsWith("Office.OneNote.", StringComparison.Ordinal)) return null;

            var isStorage = name.StartsWith(StoragePrefix, StringComparison.Ordinal);

            // ---- gather outcome evidence generically, so unknown field spellings still count ----
            bool? success = null;
            foreach (var f in SuccessFields)
                if (root.TryGetProperty(f, out var se) && se.ValueKind is JsonValueKind.True or JsonValueKind.False)
                { success = se.GetBoolean(); break; }

            uint code = 0; string? errorText = null, errorType = null; int? httpStatus = null;
            // an error field that explicitly says "No error" is POSITIVE evidence of success
            var sawExplicitNoError = false;
            var errorFields = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var p in root.EnumerateObject())
            {
                if (!p.Name.StartsWith("Data.", StringComparison.Ordinal)) continue;
                if (ErrorCodeField.IsMatch(p.Name))
                {
                    var v = AsUInt(p.Value);
                    if (v != 0) { if (code == 0) code = v; errorFields[p.Name] = v.ToString(); }
                }
                else if (HttpStatusField.IsMatch(p.Name))
                {
                    var v = AsUInt(p.Value);
                    if (v != 0) { httpStatus = (int)v; errorFields[p.Name] = v.ToString(); }
                }
                else if (ErrorTextField.IsMatch(p.Name) && p.Value.ValueKind == JsonValueKind.String)
                {
                    var v = p.Value.GetString() ?? "";
                    if (NonErrorTexts.Contains(v.Trim(), StringComparer.OrdinalIgnoreCase))
                    {
                        // "Data.Error":"No error" — only counts as success evidence, never as an error
                        if (p.Name.Equals("Data.Error", StringComparison.OrdinalIgnoreCase)) sawExplicitNoError = true;
                    }
                    else
                    {
                        errorFields[p.Name] = v;
                        if (p.Name.EndsWith("Error_Type", StringComparison.OrdinalIgnoreCase)) errorType = v;
                        // OperationWithError names the failing operation; it is a qualifier, not the message
                        else if (!p.Name.Equals("Data.OperationWithError", StringComparison.OrdinalIgnoreCase)) errorText ??= v;
                    }
                }
            }
            // Some events carry the code ONLY inside the error text, e.g. the real-time channel's
            // Data.Error = "Win32Error: ErrOutOfSyncWithStore (0xE000002E) tag_4oxx9". Without this the
            // ErrorCatalog is never consulted and the user gets a guessed category and generic advice —
            // exactly the "a general error is useless" failure mode. Field values still win.
            if (code == 0 && errorText is not null)
            {
                var m = EmbeddedHexCode.Match(errorText);
                if (m.Success && uint.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.HexNumber,
                        System.Globalization.CultureInfo.InvariantCulture, out var embedded) && embedded != 0)
                    code = embedded;
            }

            // real-time channel: "Upload" + "HTTP 503 …" → "Upload: HTTP 503 …"
            if (errorFields.TryGetValue("Data.OperationWithError", out var op))
                errorText = errorText is null ? op
                    : errorText.StartsWith(op, StringComparison.Ordinal) ? errorText : $"{op}: {errorText}";

            // An HTTP status IS an outcome, and ignoring it was a real gap: on 2026-09-08 a
            // GetUserTypesRequestFailed arrived carrying Data.HttpStatus 503 and was reported as
            // "indicates a problem, but without a definite outcome" — while the definite outcome sat
            // unread in the payload. A failing status supplies the error text when nothing else did;
            // a 2xx/3xx is positive proof the call succeeded.
            if (httpStatus >= 400) errorText ??= $"HTTP {httpStatus}";
            else if (httpStatus is >= 200 and < 400) sawExplicitNoError = true;

            var transient = root.TryGetProperty("Data.IsErrorTransient", out var tr) && tr.ValueKind == JsonValueKind.True;
            // 5xx is the server's own "try again"; 408/429 are timeout and throttling. All are retryable
            // by definition, so they escalate through [transient] rather than shouting on first sight.
            transient |= httpStatus is >= 500 or 408 or 429;
            var leaf = name[(name.LastIndexOf('.') + 1)..];
            var nameSignalsFailure = FailureWordsInName.IsMatch(leaf) && !FailureWordExceptions.IsMatch(leaf);
            var hasErrorEvidence = code != 0 || errorText is not null;

            // ---- scope: Storage.* always; other OneNote.* only when it signals failure ----
            if (!isStorage && !nameSignalsFailure && !hasErrorEvidence && success != false) return null;

            var kind = name switch
            {
                "Office.OneNote.Storage.NotebookSyncResult" => SyncEventKind.NotebookSyncResult,
                "Office.OneNote.Storage.SectionSyncResult" => SyncEventKind.SectionSyncResult,
                "Office.OneNote.Storage.SyncScore" => SyncEventKind.SyncScore,
                "Office.OneNote.Storage.PageSyncSession" => SyncEventKind.PageSyncSession,
                "Office.OneNote.Storage.ConnectivityChanged" => SyncEventKind.ConnectivityChanged,
                "Office.OneNote.Storage.RealTime.NoteItHttpUpload" => SyncEventKind.PageUpload,
                "Office.OneNote.Storage.RealTime.NoteItHttpDownload" => SyncEventKind.PageDownload,
                "Office.OneNote.Storage.RealTime.NoteItService" => SyncEventKind.RealTimeService,
                _ => isStorage ? SyncEventKind.UnrecognisedStorage : SyncEventKind.OtherOneNoteSignal,
            };

            var errorStateMs = GetLong(root, "Data.ErrorState_TimeInActiveSyncErrorState")
                             + GetLong(root, "Data.ErrorState_TimeInHierarchySyncErrorState");

            var (outcome, reason) = Classify(name, kind, success, code, errorText, transient, nameSignalsFailure, errorStateMs, sawExplicitNoError);

            return new SyncEvent
            {
                EventName = name,
                Kind = kind,
                Outcome = outcome,
                ClassificationReason = reason,
                // Fail-closed asks the user to REPORT an event it could not classify, so the payload has
                // to travel with it: while OneNote runs it holds its diagnostic log with an exclusive
                // lock, so by the time anyone reads the warning the evidence is unreachable. Kept only
                // for the events that actually need reporting, so classified traffic costs nothing.
                RawPayload = outcome is SyncOutcome.Unknown or SyncOutcome.SuspectedFailure
                    ? json.Length <= MaxRawPayload ? json : json[..MaxRawPayload] + "…(truncated)"
                    : null,
                Time = GetTime(root, "Time"),
                Success = success,
                ErrorCode = code,
                ErrorDescription = errorText,
                ErrorType = errorType,
                IsErrorTransient = transient,
                NotebookGosid = GetString(root, "Data.Gosid"),
                NotebookResourceId = GetString(root, "Data.NotebookId_ResourceId"),
                SectionResourceId = GetString(root, "Data.SectionResourceId_ResourceId") ?? GetString(root, "Data.SectionId_ResourceId"),
                SectionGosid = GetString(root, "Data.UnmappedGosid"),
                PageGoid = GetString(root, "Data.ActivePageGOID"),
                InternetAvailable = root.TryGetProperty("Data.InternetConnectivityNowAvailable", out var ia) && ia.ValueKind is JsonValueKind.True or JsonValueKind.False ? ia.GetBoolean() : null,
                TimeInSyncErrorStateMs = errorStateMs,
                TransferTimeMs = Math.Max(GetLong(root, "Data.UploadTimeInMs"), GetLong(root, "Data.TimeToConfirmSyncedWithServerInMs")),
                SyncDestinationType = GetString(root, "Data.SyncDestinationType"),
                ErrorFields = errorFields,
            };
        }
    }

    /// <summary>
    /// The fail-closed decision table. ORDER MATTERS:
    /// specific evidence → error-bearing exceptions → explicit failure → non-outcome telemetry →
    /// failure signals → success → Unknown.
    /// The <see cref="DiagnosticEvents"/> lookup sits above generic error evidence so a scan event that
    /// merely *reports* an error it skipped (SyncScore) is not mistaken for a failed sync — while a page
    /// that actually spent time in an error state is caught first and still surfaces.
    /// <see cref="BenignUnlessError"/> sits below it instead, so listing an event there silences its
    /// routine chatter without also silencing a failure it reports.
    /// </summary>
    private static (SyncOutcome, string) Classify(
        string name, SyncEventKind kind, bool? success, uint code, string? errorText,
        bool transient, bool nameSignalsFailure, long errorStateMs, bool sawExplicitNoError)
    {
        // 1. hard evidence that a real sync spent time failing
        if (errorStateMs > 0)
            return (SyncOutcome.SuspectedFailure, $"the page spent {errorStateMs} ms in a sync error state");

        // 2. deliberate, evidence-backed non-outcome telemetry
        if (DiagnosticEvents.TryGetValue(name, out var why))
            return (SyncOutcome.Diagnostic, why);

        // 3. explicit failure
        if (success == false)
            return transient
                ? (SyncOutcome.Transient, "OneNote reported failure and flagged it transient")
                : (SyncOutcome.Failure, "OneNote reported Success=false");

        if (code != 0 || errorText is not null)
        {
            var what = code != 0 ? $"error code 0x{code:X8}" : $"error text \"{errorText}\"";
            return transient
                ? (SyncOutcome.Transient, $"{what}, flagged transient by OneNote")
                : (SyncOutcome.Failure, $"OneNote reported {what}");
        }

        // 4. known non-outcome telemetry that brought no error evidence with it. Below the error checks
        //    on purpose: being on this list buys silence only for a payload that reported nothing wrong.
        if (BenignUnlessError.TryGetValue(name, out var benign))
            return (SyncOutcome.Diagnostic, benign);

        // 5. a failure signal in the name, with no outcome fields to confirm it
        if (nameSignalsFailure)
            return (SyncOutcome.SuspectedFailure, $"the event name '{name}' indicates a problem");

        // 6. positive success evidence
        if (success == true)
            return (SyncOutcome.Success, "OneNote reported success");
        if (sawExplicitNoError)
            return (SyncOutcome.Success, "OneNote explicitly reported 'No error'");
        if (kind is SyncEventKind.PageUpload or SyncEventKind.PageDownload)
            return (SyncOutcome.Success, "a page transfer completed");

        // 7. FAIL-CLOSED: no positive success signal and no known-benign classification
        return (SyncOutcome.Unknown,
            kind == SyncEventKind.UnrecognisedStorage
                ? $"'{name}' is not known to this build and carries no success flag"
                : $"'{name}' carries no success flag");
    }

    private static string? ExtractName(string json)
    {
        var m = Regex.Match(json, "\"EventName\"\\s*:\\s*\"([^\"]+)\"");
        return m.Success ? m.Groups[1].Value : null;
    }

    private static string? GetString(JsonElement root, string name) =>
        root.TryGetProperty(name, out var el) && el.ValueKind == JsonValueKind.String ? el.GetString() : null;

    private static long GetLong(JsonElement root, string name) =>
        root.TryGetProperty(name, out var el) && el.ValueKind == JsonValueKind.Number && el.TryGetInt64(out var l) ? l : 0;

    private static uint AsUInt(JsonElement el) => el.ValueKind switch
    {
        JsonValueKind.Number when el.TryGetInt64(out var l) && l is >= 0 and <= uint.MaxValue => (uint)l,
        JsonValueKind.String when uint.TryParse(el.GetString(), out var s) => s,
        _ => 0,
    };

    private static DateTimeOffset GetTime(JsonElement root, string name)
    {
        var s = GetString(root, name);
        return s is not null && DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var t) ? t : default;
    }
}
