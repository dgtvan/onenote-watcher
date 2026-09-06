using System.Runtime.InteropServices;

namespace OneNoteWatcher.Core;

/// <summary>
/// Elapsed time EXCLUDING any period the machine was asleep.
///
/// Every "nothing has happened for N minutes" check in this app is really asking "have I been watching
/// for N minutes and seen nothing?". Wall-clock subtraction answers a different question, because a
/// suspended machine produces a gap the watcher was never awake to observe. On 2026-09-06 a 3-hour
/// sleep raised two false alarms at once — "no Office telemetry for 3.0 h" and "the Microsoft Graph
/// check is not running" — neither of which described anything wrong.
///
/// <c>QueryUnbiasedInterruptTime</c> is the OS counter that stops while the system is suspended, which
/// is exactly the semantics these checks need. Take a <see cref="Stamp"/> when something happens and
/// measure with <see cref="Since"/>.
/// </summary>
public static class AwakeClock
{
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool QueryUnbiasedInterruptTime(out ulong unbiasedTime);

    /// <summary>True when the OS counter is usable; false on a platform where we fall back to wall clock.</summary>
    public static bool Available { get; } = TryRead(out _);

    private static bool TryRead(out ulong t)
    {
        try { return QueryUnbiasedInterruptTime(out t); }
        catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException)
        { t = 0; return false; }
    }

    /// <summary>A stamp to compare later. Units are opaque; only differences are meaningful.</summary>
    public static TimeSpan Stamp() =>
        TryRead(out var t) ? TimeSpan.FromTicks((long)(t & long.MaxValue)) : TimeSpan.FromTicks(DateTime.UtcNow.Ticks);

    /// <summary>Awake time elapsed since <paramref name="stamp"/>, never negative.</summary>
    public static TimeSpan Since(TimeSpan stamp)
    {
        var d = Stamp() - stamp;
        return d < TimeSpan.Zero ? TimeSpan.Zero : d;
    }

    /// <summary>
    /// The smaller of the wall-clock gap and the awake gap. Used where a wall-clock timestamp is the
    /// only thing available for the "last seen" side: it can never over-report silence, so a sleep can
    /// never manufacture an alarm, while a genuine stall is still reported in full.
    /// </summary>
    public static TimeSpan Elapsed(DateTimeOffset wallSince, DateTimeOffset wallNow, TimeSpan awakeStamp)
    {
        var wall = wallNow - wallSince;
        if (wall < TimeSpan.Zero) wall = TimeSpan.Zero;
        var awake = Since(awakeStamp);
        return awake < wall ? awake : wall;
    }
}
