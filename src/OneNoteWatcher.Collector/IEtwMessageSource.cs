namespace OneNoteWatcher.Collector;

/// <summary>One OfficeLoggingLiblet "etwtaskLogging" record reduced to what the detector needs.</summary>
public readonly record struct OfficeLogMessage(DateTime TimeUtc, string Category, string Message);

/// <summary>
/// Yields OneNote's OfficeLoggingLiblet telemetry messages. Two implementations: a live
/// <see cref="LiveEtwMessageSource"/> (elevated real-time session) and an
/// <see cref="EtlFileMessageSource"/> (offline replay of a captured .etl — no admin, for tests).
/// </summary>
public interface IEtwMessageSource
{
    /// <summary>Blocks, invoking <paramref name="onMessage"/> per record until cancelled/EOF.</summary>
    void Process(Action<OfficeLogMessage> onMessage, CancellationToken ct);
}
