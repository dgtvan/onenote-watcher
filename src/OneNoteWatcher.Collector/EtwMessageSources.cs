using Microsoft.Diagnostics.Tracing;
using Microsoft.Diagnostics.Tracing.Session;

namespace OneNoteWatcher.Collector;

/// <summary>
/// Provider constants for OneNote/Office logging over ETW. Verified in docs/detection-design.md.
/// </summary>
public static class OfficeEtw
{
    public static readonly Guid OfficeLoggingLiblet = new("F50D9315-E17E-43C1-8370-3EDF6CC057BE");
    public const string LoggingEventName = "etwtaskLogging";
    public const string TelemetryCategory = "Telemetry Event";
    public const string MessageField = "wzMessage";
    public const string CategoryField = "wzCategory";
    public const string ProductField = "wzProduct";
    public const string OneNoteProduct = "Microsoft OneNote";
}

/// <summary>Live real-time ETW session on OfficeLoggingLiblet. Requires elevation to start.</summary>
public sealed class LiveEtwMessageSource : IEtwMessageSource
{
    private readonly string _sessionName;
    /// <summary>Raised with the number of ETW records Windows dropped, so a coverage gap is never silent.</summary>
    public Action<int>? OnEventsLost { get; set; }

    public LiveEtwMessageSource(string sessionName = "OneNoteWatcher_Sync") => _sessionName = sessionName;

    public void Process(Action<OfficeLogMessage> onMessage, CancellationToken ct)
    {
        // A restarted session with the same name replaces any leftover from a crash.
        using var session = new TraceEventSession(_sessionName);
        using var reg = ct.Register(() => { try { session.Stop(); } catch { } });

        session.EnableProvider(OfficeEtw.OfficeLoggingLiblet, TraceEventLevel.Verbose);

        session.Source.Dynamic.All += e =>
        {
            if (e.ProviderGuid != OfficeEtw.OfficeLoggingLiblet) return;
            if (!string.Equals(e.EventName, OfficeEtw.LoggingEventName, StringComparison.Ordinal)) return;

            var category = e.PayloadStringByName(OfficeEtw.CategoryField);
            if (!string.Equals(category, OfficeEtw.TelemetryCategory, StringComparison.Ordinal)) return;

            var message = e.PayloadStringByName(OfficeEtw.MessageField);
            if (message is null) return;

            onMessage(new OfficeLogMessage(e.TimeStamp.ToUniversalTime(), category, message));
        };

        var lastLost = 0;
        session.Source.Process(); // blocks until Stop()
        var lost = session.Source.EventsLost;
        if (lost > lastLost) OnEventsLost?.Invoke(lost - lastLost);
    }
}

/// <summary>Offline replay of a captured .etl. No elevation needed — used to verify the detector.</summary>
public sealed class EtlFileMessageSource : IEtwMessageSource
{
    private readonly string _etlPath;
    public EtlFileMessageSource(string etlPath) => _etlPath = etlPath;

    public void Process(Action<OfficeLogMessage> onMessage, CancellationToken ct)
    {
        using var source = new ETWTraceEventSource(_etlPath);
        source.Dynamic.All += e =>
        {
            if (ct.IsCancellationRequested) { source.StopProcessing(); return; }
            if (e.ProviderGuid != OfficeEtw.OfficeLoggingLiblet) return;
            if (!string.Equals(e.EventName, OfficeEtw.LoggingEventName, StringComparison.Ordinal)) return;

            var category = e.PayloadStringByName(OfficeEtw.CategoryField);
            if (!string.Equals(category, OfficeEtw.TelemetryCategory, StringComparison.Ordinal)) return;

            var message = e.PayloadStringByName(OfficeEtw.MessageField);
            if (message is null) return;

            onMessage(new OfficeLogMessage(e.TimeStamp.ToUniversalTime(), category, message));
        };
        source.Process();
    }
}
