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

    /// <summary>Words in an event NAME that mean "something went wrong" even with no error field.</summary>
    private static readonly Regex FailureWordsInName = new(
        "Fail|Error|Blocker|Blocked|Stuck|Corrupt|ReadOnly|Conflict|Denied|Unauthori|Expired|Rejected|Abort|Crash|Inconsistenc",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>Field names that carry an explicit success flag.</summary>
    private static readonly string[] SuccessFields = ["Data.Success", "Data.IsSuccess", "Data.Succeeded", "Data.WasSuccessful"];

    /// <summary>Field-name pattern for numeric error codes (Error_Code, NotebookErrorCode, SH_ErrorCode, …).</summary>
    private static readonly Regex ErrorCodeField = new(@"^Data\.[A-Za-z0-9_.]*Error_?Code$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>An 8-digit hex code embedded in free-text error output, e.g. "(0xE000002E)".</summary>
    private static readonly Regex EmbeddedHexCode = new(@"0x([0-9A-Fa-f]{8})(?![0-9A-Fa-f])", RegexOptions.Compiled);

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

            uint code = 0; string? errorText = null, errorType = null;
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

            var transient = root.TryGetProperty("Data.IsErrorTransient", out var tr) && tr.ValueKind == JsonValueKind.True;
            var nameSignalsFailure = FailureWordsInName.IsMatch(name[(name.LastIndexOf('.') + 1)..]);
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
    /// specific evidence → documented exceptions → explicit failure → failure signals → success → Unknown.
    /// The exception lookup sits above generic error evidence so a scan event that merely *reports* an
    /// error it skipped (SyncScore) is not mistaken for a failed sync — while a page that actually spent
    /// time in an error state is caught first and still surfaces.
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

        // 4. a failure signal in the name, with no outcome fields to confirm it
        if (nameSignalsFailure)
            return (SyncOutcome.SuspectedFailure, $"the event name '{name}' indicates a problem");

        // 5. positive success evidence
        if (success == true)
            return (SyncOutcome.Success, "OneNote reported success");
        if (sawExplicitNoError)
            return (SyncOutcome.Success, "OneNote explicitly reported 'No error'");
        if (kind is SyncEventKind.PageUpload or SyncEventKind.PageDownload)
            return (SyncOutcome.Success, "a page transfer completed");

        // 6. FAIL-CLOSED: no positive success signal and no known-benign classification
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
