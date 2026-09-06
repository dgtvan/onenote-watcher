namespace OneNoteWatcher.Core.Model;

/// <summary>
/// How a sync event was classified. The watcher is FAIL-CLOSED: only an explicit, positive success
/// signal yields <see cref="Success"/>; anything we cannot prove successful is surfaced.
/// </summary>
public enum SyncOutcome
{
    /// <summary>Explicit success (Data.Success/IsSuccess true, or a completed transfer) with no error field.</summary>
    Success,
    /// <summary>Explicit failure: success flag false, or a non-zero error code, or an error message.</summary>
    Failure,
    /// <summary>OneNote flagged the error transient (retryable). Warn; escalates on repeat via [transient].</summary>
    Transient,
    /// <summary>A failure *signal* without a confirmed outcome (e.g. an event named "…SyncBlocker…").</summary>
    SuspectedFailure,
    /// <summary>An event we do not recognise and cannot prove successful. Fail-closed → surfaced.</summary>
    Unknown,
    /// <summary>Deliberately classified as non-outcome telemetry (metrics, scans). Recorded, never alerted.</summary>
    Diagnostic,
}

/// <summary>
/// A OneNote telemetry record parsed from a <c>SendEvent {…}</c> payload — the same shape live over
/// ETW (OfficeLoggingLiblet <c>wzMessage</c>) and at rest in the Office diagnostic log.
/// See docs/detection-design.md and docs/fail-closed.md.
/// </summary>
public sealed record SyncEvent
{
    public required string EventName { get; init; }
    public SyncEventKind Kind { get; init; }
    public required SyncOutcome Outcome { get; init; }

    /// <summary>UTC timestamp from the payload "Time" field.</summary>
    public DateTimeOffset Time { get; init; }

    /// <summary>The explicit success flag when the payload carried one (Data.Success / Data.IsSuccess).</summary>
    public bool? Success { get; init; }

    /// <summary>Unsigned 32-bit OneNote error code (0 = none), from any *ErrorCode / Error_Code field.</summary>
    public uint ErrorCode { get; init; }

    /// <summary>Symbolic description, error text, or the reason this event was classified as it was.</summary>
    public string? ErrorDescription { get; init; }
    public string? ErrorType { get; init; }
    public bool IsErrorTransient { get; init; }

    /// <summary>Why the classifier reached <see cref="Outcome"/> — shown in logs and issue details.</summary>
    public string? ClassificationReason { get; init; }

    public string? NotebookGosid { get; init; }
    public string? NotebookResourceId { get; init; }
    public string? SectionResourceId { get; init; }
    public string? SectionGosid { get; init; }
    public string? PageGoid { get; init; }

    /// <summary>For ConnectivityChanged: Data.InternetConnectivityNowAvailable.</summary>
    public bool? InternetAvailable { get; init; }

    /// <summary>PageSyncSession: milliseconds the active page spent in an error state.</summary>
    public long TimeInSyncErrorStateMs { get; init; }

    /// <summary>Real-time page upload/download duration.</summary>
    public long TransferTimeMs { get; init; }

    public string? SyncDestinationType { get; init; }

    /// <summary>Error-bearing raw fields kept verbatim for the issue's technical detail.</summary>
    public IReadOnlyDictionary<string, string> ErrorFields { get; init; } =
        new Dictionary<string, string>();

    /// <summary>Anything that must reach the user as a problem (alert or warning).</summary>
    public bool IsProblem => Outcome is SyncOutcome.Failure or SyncOutcome.Transient
        or SyncOutcome.SuspectedFailure or SyncOutcome.Unknown;

    /// <summary>True for events that prove content moved between this PC and OneDrive.</summary>
    public bool IsSyncActivity =>
        Outcome == SyncOutcome.Success
        && Kind is SyncEventKind.NotebookSyncResult or SyncEventKind.SectionSyncResult
                or SyncEventKind.PageUpload or SyncEventKind.PageDownload;

    public string? ErrorCodeHex => ErrorCode == 0 ? null : $"0x{ErrorCode:X8}";
}

public enum SyncEventKind
{
    Unknown = 0,
    NotebookSyncResult,
    SectionSyncResult,
    SyncScore,
    PageSyncSession,
    ConnectivityChanged,
    PageUpload,
    PageDownload,
    RealTimeService,
    /// <summary>An <c>Office.OneNote.Storage.*</c> event this build does not know. Fail-closed.</summary>
    UnrecognisedStorage,
    /// <summary>A non-Storage <c>Office.OneNote.*</c> event carrying a failure signal.</summary>
    OtherOneNoteSignal,
}
