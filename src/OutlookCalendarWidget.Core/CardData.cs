using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OutlookCalendarWidget.Core;

/// <summary>Well-known launch URIs handled by the companion app (declared as a protocol in the package manifest).</summary>
public static class AppLinks
{
    public const string Scheme = "outlook-calendar-widget";
    public const string SignIn = Scheme + ":signin";
    public const string Settings = Scheme + ":settings";
    public const string DefaultOutlookCalendarUrl = "https://outlook.office.com/calendar/view/day";

    /// <summary>Brings classic Outlook to the front on its calendar (starting it if needed).</summary>
    public const string OpenOutlook = Scheme + ":open";

    private const string StartFormat = "yyyyMMdd'T'HHmmss'Z'";

    /// <summary>Opens one calendar item (or, for a recurring series, the occurrence at <paramref name="start"/>) in classic Outlook.</summary>
    public static string OpenOutlookItem(string entryId, DateTimeOffset start) =>
        $"{OpenOutlook}?id={Uri.EscapeDataString(entryId)}&start={start.UtcDateTime.ToString(StartFormat, CultureInfo.InvariantCulture)}";

    /// <summary>Parses an <see cref="OpenOutlook"/> link. Returns false for any other URI.</summary>
    public static bool TryParseOpenOutlook(string? uri, out OutlookOpenRequest request)
    {
        request = new OutlookOpenRequest(null, null);
        if (uri is null || !uri.StartsWith(OpenOutlook, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var rest = uri[OpenOutlook.Length..];
        if (rest.Length > 0 && rest[0] != '?' && rest[0] != '/')
        {
            return false;
        }

        string? entryId = null;
        DateTimeOffset? start = null;
        foreach (var pair in rest.TrimStart('/', '?').Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = pair.Split('=', 2);
            var value = parts.Length == 2 ? Uri.UnescapeDataString(parts[1]) : string.Empty;
            if (parts[0].Equals("id", StringComparison.OrdinalIgnoreCase) && value.Length > 0)
            {
                entryId = value;
            }
            else if (parts[0].Equals("start", StringComparison.OrdinalIgnoreCase) &&
                     DateTimeOffset.TryParseExact(value, StartFormat, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var parsed))
            {
                start = parsed;
            }
        }

        request = new OutlookOpenRequest(entryId, start);
        return true;
    }
}

/// <summary>What an <see cref="AppLinks.OpenOutlook"/> link asks for: a specific item, or just the calendar.</summary>
public sealed record OutlookOpenRequest(string? EntryId, DateTimeOffset? Start);

/// <summary>Widget card states. Serialized as the <c>state</c> field and used by template <c>$when</c> conditions.</summary>
public static class CardStates
{
    public const string Ok = "ok";
    public const string Loading = "loading";
    public const string SignedOut = "signedOut";
    public const string NotConfigured = "notConfigured";
    public const string Unavailable = "unavailable";
    public const string Error = "error";
}

/// <summary>The data document bound to <c>CalendarWidget.json</c> by the Widgets Board's Adaptive Card templating engine.</summary>
public sealed record CardData
{
    public string State { get; init; } = CardStates.Ok;

    public bool IsOk => State == CardStates.Ok;

    public string StatusTitle { get; init; } = string.Empty;

    public string Message { get; init; } = string.Empty;

    public string ActionTitle { get; init; } = string.Empty;

    public string ActionUrl { get; init; } = string.Empty;

    public string ActionVerb { get; init; } = string.Empty;

    public bool HasActionUrl => ActionUrl.Length > 0;

    public bool HasActionVerb => ActionUrl.Length == 0 && ActionVerb.Length > 0;

    public string Heading { get; init; } = string.Empty;

    public string DateLabel { get; init; } = string.Empty;

    public string Summary { get; init; } = string.Empty;

    public bool HasNext => Next is not null;

    public CardEvent? Next { get; init; }

    public bool HasEvents => Days.Count > 0;

    public bool ShowDayHeaders { get; init; }

    public IReadOnlyList<CardDay> Days { get; init; } = [];

    public string EmptyTitle { get; init; } = string.Empty;

    public string EmptyText { get; init; } = string.Empty;

    public bool HasMore => MoreText.Length > 0;

    public string MoreText { get; init; } = string.Empty;

    public string UpdatedText { get; init; } = string.Empty;

    public string OutlookUrl { get; init; } = AppLinks.DefaultOutlookCalendarUrl;

    public string ToJson() => JsonSerializer.Serialize(this, CardJsonContext.Default.CardData);

    public static CardData ForStatus(
        string state,
        string title,
        string message,
        string actionTitle = "",
        string actionUrl = "",
        string actionVerb = "",
        string updatedText = "") => new()
        {
            State = state,
            StatusTitle = title,
            Message = message,
            ActionTitle = actionTitle,
            ActionUrl = actionUrl,
            ActionVerb = actionVerb,
            UpdatedText = updatedText,
        };
}

public sealed record CardDay
{
    public required string Label { get; init; }

    public required IReadOnlyList<CardEvent> Events { get; init; }
}

public sealed record CardEvent
{
    public required string Subject { get; init; }

    public required string TimeText { get; init; }

    public string EndText { get; init; } = string.Empty;

    public bool HasEndText => EndText.Length > 0;

    public string RelativeText { get; init; } = string.Empty;

    public string Location { get; init; } = string.Empty;

    public bool HasLocation => Location.Length > 0;

    /// <summary>Adaptive Card container style used for the free/busy color bar.</summary>
    public string StatusStyle { get; init; } = "accent";

    /// <summary>PNG data URI of the colored free/busy bar for <see cref="StatusStyle"/>.</summary>
    public string StatusBarUrl => StatusBar.DataUri(StatusStyle);

    public string StatusText { get; init; } = string.Empty;

    public bool HasStatusText => StatusText.Length > 0;

    public string Detail { get; init; } = string.Empty;

    public bool HasDetail => Detail.Length > 0;

    /// <summary>Opens the event in Outlook (falls back to the Outlook calendar when Graph returned no link).</summary>
    public string OpenUrl { get; init; } = AppLinks.DefaultOutlookCalendarUrl;

    public string JoinUrl { get; init; } = string.Empty;

    public bool HasJoin => JoinUrl.Length > 0;

    /// <summary>True when the event has an online meeting that is in progress or starts soon.</summary>
    public bool ShowJoin { get; init; }

    /// <summary>The single most useful tap target: join when the meeting is imminent, otherwise open it.</summary>
    public string PrimaryUrl => ShowJoin ? JoinUrl : OpenUrl;

    public bool IsNow { get; init; }

    public bool IsAllDay { get; init; }

    public string RowStyle => IsNow ? "emphasis" : "default";
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(CardData))]
internal sealed partial class CardJsonContext : JsonSerializerContext;
