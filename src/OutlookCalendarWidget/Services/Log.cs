using System.Diagnostics;

namespace OutlookCalendarWidget.Services;

/// <summary>
/// Minimal append-only diagnostics log in the package's LocalState folder. The widget provider runs headless inside
/// a COM server process, so this is the easiest way to see why a widget is misbehaving.
/// </summary>
internal static class Log
{
    private const long MaxBytes = 512 * 1024;
    private static readonly Lock FileLock = new();
    private static readonly string Role = Environment.GetCommandLineArgs().Any(a => a.Contains("ComServer", StringComparison.OrdinalIgnoreCase)) ? "provider" : "app";

    public static string FilePath => Path.Combine(AppSettings.DataFolder, "logs", "diagnostics.log");

    public static void Info(string message) => Write("INFO", message);

    public static void Error(string message, Exception exception) => Write("ERROR", $"{message}: {exception}");

    private static void Write(string level, string message)
    {
        var line = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff zzz} [{Role}:{Environment.ProcessId}] {level} {message}{Environment.NewLine}";
        Debug.Write(line);
        try
        {
            lock (FileLock)
            {
                var path = FilePath;
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                if (File.Exists(path) && new FileInfo(path).Length > MaxBytes)
                {
                    File.Move(path, path + ".old", overwrite: true);
                }

                File.AppendAllText(path, line);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Another process may hold the file briefly; diagnostics must never break the widget.
        }
    }
}
