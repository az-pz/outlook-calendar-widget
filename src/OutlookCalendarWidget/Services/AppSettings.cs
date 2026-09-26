using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using OutlookCalendarWidget.Core;
using Windows.Storage;

namespace OutlookCalendarWidget.Services;

/// <summary>Where the widget gets its events from.</summary>
internal enum CalendarSource
{
    /// <summary>Classic Outlook for Windows running on this PC (COM object model). No sign-in needed.</summary>
    Outlook,

    /// <summary>Microsoft Graph, signed in through the Windows account broker. Needs an Entra app registration.</summary>
    Microsoft365,

    /// <summary><see cref="SampleCalendar"/> events (demo mode).</summary>
    Sample,
}

/// <summary>
/// User settings shared by the UI process and the widget provider process (both run with the same package identity,
/// so they share the package's LocalState folder).
/// </summary>
internal sealed record AppSettings
{
    private static readonly Lock FileLock = new();

    [JsonConverter(typeof(JsonStringEnumConverter<CalendarSource>))]
    public CalendarSource Source { get; init; } = CalendarSource.Outlook;

    /// <summary>Application (client) ID of the Entra ID app registration used to call Microsoft Graph.</summary>
    public string ClientId { get; init; } = BuildDefault("GraphClientId") ?? string.Empty;

    /// <summary>Tenant: <c>common</c>, <c>organizations</c>, <c>consumers</c>, a tenant ID, or a verified domain.</summary>
    public string TenantId { get; init; } = BuildDefault("GraphTenantId") ?? "common";

    /// <summary>Where "Open Outlook" goes when the source is Microsoft Graph (the Outlook source opens the desktop app).</summary>
    public string OutlookUrl { get; init; } = AppLinks.DefaultOutlookCalendarUrl;

    /// <summary>Set after an explicit sign-out so the widget stops using the Windows account silently.</summary>
    public bool SignedOut { get; init; }

    [JsonIgnore]
    public bool IsConfigured => Guid.TryParse(ClientId, out _);

    public static string DataFolder => ApplicationData.Current.LocalFolder.Path;

    private static string FilePath => Path.Combine(DataFolder, "settings.json");

    public static AppSettings Load()
    {
        lock (FileLock)
        {
            try
            {
                if (File.Exists(FilePath))
                {
                    var settings = JsonSerializer.Deserialize(File.ReadAllText(FilePath), SettingsJsonContext.Default.AppSettings);
                    if (settings is not null)
                    {
                        return settings;
                    }
                }
            }
            catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
            {
            }

            return new AppSettings();
        }
    }

    public void Save()
    {
        lock (FileLock)
        {
            Directory.CreateDirectory(DataFolder);
            var temp = FilePath + ".tmp";
            File.WriteAllText(temp, JsonSerializer.Serialize(this, SettingsJsonContext.Default.AppSettings));
            File.Move(temp, FilePath, overwrite: true);
        }
    }

    private static string? BuildDefault(string key)
    {
        var value = typeof(AppSettings).Assembly
            .GetCustomAttributes<AssemblyMetadataAttribute>()
            .FirstOrDefault(a => a.Key == key)?.Value;
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }
}

[JsonSourceGenerationOptions(WriteIndented = true, PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(AppSettings))]
internal sealed partial class SettingsJsonContext : JsonSerializerContext;
