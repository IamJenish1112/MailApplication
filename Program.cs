using MailApplication.Forms;
using MongoDB.Bson;
using MongoDB.Driver;

namespace MailApplication;

internal static class Program
{
    private static readonly string LogFile = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "startup_error.log");

    [STAThread]
    static void Main()
    {
        // ── Global unhandled exception handlers ────────────────
        Application.ThreadException += (s, e) =>
        {
            LogAndShow("UI Thread Exception", e.Exception);
        };

        AppDomain.CurrentDomain.UnhandledException += (s, e) =>
        {
            if (e.ExceptionObject is Exception ex)
                LogAndShow("Fatal Exception", ex);
            else
                LogCrash($"Unknown fatal error: {e.ExceptionObject}");
        };

        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

        // ── Run diagnostic checks before starting ─────────────
        var (ok, message) = RunStartupDiagnostics();
        if (!ok)
        {
            MessageBox.Show(message,
                "Startup Problem Detected",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            // Still continue — warnings are not fatal
        }

        try
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
        catch (Exception ex)
        {
            LogAndShow("Startup Crash", ex);
        }
    }

    /// <summary>
    /// Checks common runtime requirements and returns a warning if anything is missing.
    /// </summary>
    private static (bool ok, string message) RunStartupDiagnostics()
    {
        var issues = new List<string>();

        // 1. Check MongoDB connectivity
        try
        {
            var client = new MongoClient("mongodb://localhost:27017");
            var db = client.GetDatabase("admin");
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(3));
            db.RunCommand<BsonDocument>(new BsonDocument("ping", 1), cancellationToken: cts.Token);
        }
        catch (OperationCanceledException)
        {
            issues.Add("⚠ MongoDB is not running on localhost:27017.\n   → Start MongoDB before using this app.");
        }
        catch (Exception ex)
        {
            issues.Add($"⚠ MongoDB connection error: {ex.Message}");
        }

        // 2. Check WebView2 Runtime
        try
        {
            var wv2Version = Microsoft.Web.WebView2.Core.CoreWebView2Environment.GetAvailableBrowserVersionString();
            if (string.IsNullOrEmpty(wv2Version))
                issues.Add("⚠ WebView2 Runtime not found.\n   → Download from: https://developer.microsoft.com/en-us/microsoft-edge/webview2/");
        }
        catch
        {
            issues.Add("⚠ WebView2 Runtime is not installed.\n   → Download from: https://developer.microsoft.com/en-us/microsoft-edge/webview2/\n   → The email editor will NOT work without it.");
        }

        // 3. Check editor HTML exists
        var editorPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "editor", "editor.html");
        if (!File.Exists(editorPath))
        {
            issues.Add($"⚠ Editor file not found at:\n   {editorPath}\n   → Rebuild the project to copy wwwroot files.");
        }

        if (issues.Count == 0)
            return (true, string.Empty);

        var msg = string.Join("\n\n", issues);
        LogCrash($"Startup diagnostics found issues:\n{msg}");
        return (false, msg);
    }

    private static void LogAndShow(string title, Exception ex)
    {
        var details = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {title}\n" +
                      $"Message: {ex.Message}\n" +
                      $"Type:    {ex.GetType().FullName}\n" +
                      $"Source:  {ex.Source}\n" +
                      $"Stack:\n{ex.StackTrace}" +
                      (ex.InnerException != null ? $"\n\nInner: {ex.InnerException.Message}\n{ex.InnerException.StackTrace}" : "");

        LogCrash(details);

        MessageBox.Show(
            $"An unexpected error occurred:\n\n" +
            $"{ex.Message}\n\n" +
            $"Type: {ex.GetType().Name}\n\n" +
            $"Details saved to:\n{LogFile}",
            title,
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
    }

    private static void LogCrash(string message)
    {
        try { File.AppendAllText(LogFile, message + "\n" + new string('-', 60) + "\n"); }
        catch { /* never throw from logging */ }
    }
}
