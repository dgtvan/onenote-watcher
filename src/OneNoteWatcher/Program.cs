using System.Diagnostics;

namespace OneNoteWatcher;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Contains("--graph-check"))
        {
            // one-shot diagnostic report; runs without the tray (and without the single-instance mutex)
            var path = GraphCheck.RunAsync(Path.Combine(AppContext.BaseDirectory, "config.ini")).GetAwaiter().GetResult();
            if (!args.Contains("--quiet"))
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            return 0;
        }

        using var mutex = new Mutex(initiallyOwned: true, "OneNoteWatcher.Tray.SingleInstance", out var isNew);
        if (!isNew) return 1;

        ApplicationConfiguration.Initialize();

        // a crash must leave evidence in the log, never vanish the icon silently
        var crashLog = new OneNoteWatcher.Core.Logging.AppLog(
            Path.Combine(@"C:\ProgramData\OneNoteWatcher", "logs"), "tray");
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
            crashLog.Error("UNHANDLED EXCEPTION — tray is stopping", e.ExceptionObject as Exception);
        Application.ThreadException += (_, e) =>
            crashLog.Error("UNHANDLED UI EXCEPTION", e.Exception);
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        bool simulate = args.Contains("--simulate") || args.Contains("--selftest");
        using var app = new TrayApp(simulate);
        Application.Run();
        GC.KeepAlive(mutex);
        return 0;
    }
}
