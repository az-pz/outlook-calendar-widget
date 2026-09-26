using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OutlookCalendarWidget.Core;

/// <summary>Thrown when Microsoft Graph rejects the access token and the user must sign in again.</summary>
public sealed class GraphAuthenticationException(string message) : Exception(message);

/// <summary>Thrown when Microsoft Graph returns a non-success status code.</summary>
public sealed class GraphRequestException(HttpStatusCode statusCode, string message) : Exception(message)
{
    public HttpStatusCode StatusCode { get; } = statusCode;
}

/// <summary>Reads the signed-in user's calendar through the Microsoft Graph <c>calendarView</c> API.</summary>
public sealed class GraphCalendarClient(HttpClient httpClient, Func<CancellationToken, Task<string>> accessTokenProvider)
{
    public const string GraphScope = "Calendars.Read";

    private const string CalendarViewEndpoint = "https://graph.microsoft.com/v1.0/me/calendarView";
    private const string SelectedFields =
        "id,subject,start,end,isAllDay,isCancelled,showAs,location,responseStatus,onlineMeeting,onlineMeetingUrl,webLink";
    private const int MaxPages = 5;

    public async Task<IReadOnlyList<CalendarEvent>> GetEventsAsync(
        DateTimeOffset rangeStart,
        DateTimeOffset rangeEnd,
        TimeZoneInfo timeZone,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(timeZone);

        var url = BuildCalendarViewUrl(rangeStart, rangeEnd);
        var token = await accessTokenProvider(cancellationToken);
        var events = new List<CalendarEvent>();

        for (var page = 0; page < MaxPages && url is not null; page++)
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            request.Headers.Add("Prefer", $"outlook.timezone=\"{timeZone.Id}\"");
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            using var response = await httpClient.SendAsync(request, cancellationToken);
            var body = await response.Content.ReadAsStringAsync(cancellationToken);

            if (response.StatusCode == HttpStatusCode.Unauthorized)
            {
                throw new GraphAuthenticationException("Your session expired. Sign in again.");
            }

            if (!response.IsSuccessStatusCode)
            {
                throw new GraphRequestException(response.StatusCode, GraphCalendarParser.ReadErrorMessage(body, response.StatusCode));
            }

            var (pageEvents, nextLink) = GraphCalendarParser.ParsePage(body, timeZone);
            events.AddRange(pageEvents);
            url = nextLink;
        }

        return events;
    }

    internal static string BuildCalendarViewUrl(DateTimeOffset rangeStart, DateTimeOffset rangeEnd)
    {
        static string Format(DateTimeOffset value) =>
            Uri.EscapeDataString(value.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture));

        return $"{CalendarViewEndpoint}?startDateTime={Format(rangeStart)}&endDateTime={Format(rangeEnd)}" +
               $"&$select={SelectedFields}&$orderby=start/dateTime&$top=100";
    }
}

/// <summary>Converts Microsoft Graph event payloads into <see cref="CalendarEvent"/> instances.</summary>
public static class GraphCalendarParser
{
    public static (IReadOnlyList<CalendarEvent> Events, string? NextLink) ParsePage(string json, TimeZoneInfo localTimeZone)
    {
        ArgumentNullException.ThrowIfNull(localTimeZone);

        var page = JsonSerializer.Deserialize(json, GraphJsonContext.Default.GraphEventPage)
                   ?? throw new JsonException("Microsoft Graph returned an empty response.");

        var events = new List<CalendarEvent>(page.Value.Count);
        foreach (var item in page.Value)
        {
            if (item.Start is null || item.End is null)
            {
                continue;
            }

            events.Add(new CalendarEvent
            {
                Id = item.Id ?? string.Empty,
                Subject = string.IsNullOrWhiteSpace(item.Subject) ? string.Empty : item.Subject.Trim(),
                Start = ToLocal(item.Start, localTimeZone),
                End = ToLocal(item.End, localTimeZone),
                IsAllDay = item.IsAllDay,
                IsCancelled = item.IsCancelled,
                Location = string.IsNullOrWhiteSpace(item.Location?.DisplayName) ? null : item.Location!.DisplayName!.Trim(),
                ShowAs = ParseShowAs(item.ShowAs),
                Response = ParseResponse(item.ResponseStatus?.Response),
                JoinUrl = NullIfEmpty(item.OnlineMeeting?.JoinUrl) ?? NullIfEmpty(item.OnlineMeetingUrl),
                WebLink = NullIfEmpty(item.WebLink),
            });
        }

        return (events, NullIfEmpty(page.NextLink));
    }

    internal static string ReadErrorMessage(string body, HttpStatusCode statusCode)
    {
        try
        {
            var error = JsonSerializer.Deserialize(body, GraphJsonContext.Default.GraphErrorResponse);
            if (!string.IsNullOrWhiteSpace(error?.Error?.Message))
            {
                return error.Error.Message;
            }
        }
        catch (JsonException)
        {
        }

        return $"Microsoft Graph returned {(int)statusCode} ({statusCode}).";
    }

    internal static DateTimeOffset ToLocal(GraphDateTimeTimeZone value, TimeZoneInfo localTimeZone)
    {
        var wallClock = DateTime.Parse(
            value.DateTime ?? throw new JsonException("Event is missing a dateTime value."),
            CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces);
        wallClock = DateTime.SpecifyKind(wallClock, DateTimeKind.Unspecified);

        var sourceZone = ResolveTimeZone(value.TimeZone, localTimeZone);
        var offset = sourceZone.IsInvalidTime(wallClock)
            ? sourceZone.GetUtcOffset(wallClock.AddHours(1))
            : sourceZone.GetUtcOffset(wallClock);

        return TimeZoneInfo.ConvertTime(new DateTimeOffset(wallClock, offset), localTimeZone);
    }

    private static TimeZoneInfo ResolveTimeZone(string? id, TimeZoneInfo fallback)
    {
        if (string.IsNullOrWhiteSpace(id) || string.Equals(id, fallback.Id, StringComparison.OrdinalIgnoreCase))
        {
            return fallback;
        }

        if (string.Equals(id, "UTC", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(id, "Etc/UTC", StringComparison.OrdinalIgnoreCase))
        {
            return TimeZoneInfo.Utc;
        }

        return TimeZoneInfo.TryFindSystemTimeZoneById(id, out var zone) ? zone : fallback;
    }

    private static FreeBusyStatus ParseShowAs(string? value) => value?.ToLowerInvariant() switch
    {
        "free" => FreeBusyStatus.Free,
        "tentative" => FreeBusyStatus.Tentative,
        "busy" => FreeBusyStatus.Busy,
        "oof" => FreeBusyStatus.Oof,
        "workingelsewhere" => FreeBusyStatus.WorkingElsewhere,
        _ => FreeBusyStatus.Unknown,
    };

    private static MeetingResponse ParseResponse(string? value) => value?.ToLowerInvariant() switch
    {
        "organizer" => MeetingResponse.Organizer,
        "tentativelyaccepted" => MeetingResponse.TentativelyAccepted,
        "accepted" => MeetingResponse.Accepted,
        "declined" => MeetingResponse.Declined,
        "notresponded" => MeetingResponse.NotResponded,
        _ => MeetingResponse.None,
    };

    private static string? NullIfEmpty(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;
}

internal sealed class GraphEventPage
{
    public List<GraphEvent> Value { get; set; } = [];

    [JsonPropertyName("@odata.nextLink")]
    public string? NextLink { get; set; }
}

internal sealed class GraphEvent
{
    public string? Id { get; set; }

    public string? Subject { get; set; }

    public GraphDateTimeTimeZone? Start { get; set; }

    public GraphDateTimeTimeZone? End { get; set; }

    public bool IsAllDay { get; set; }

    public bool IsCancelled { get; set; }

    public string? ShowAs { get; set; }

    public GraphLocation? Location { get; set; }

    public GraphResponseStatus? ResponseStatus { get; set; }

    public GraphOnlineMeeting? OnlineMeeting { get; set; }

    public string? OnlineMeetingUrl { get; set; }

    public string? WebLink { get; set; }
}

internal sealed class GraphDateTimeTimeZone
{
    public string? DateTime { get; set; }

    public string? TimeZone { get; set; }
}

internal sealed class GraphLocation
{
    public string? DisplayName { get; set; }
}

internal sealed class GraphResponseStatus
{
    public string? Response { get; set; }
}

internal sealed class GraphOnlineMeeting
{
    public string? JoinUrl { get; set; }
}

internal sealed class GraphErrorResponse
{
    public GraphError? Error { get; set; }
}

internal sealed class GraphError
{
    public string? Code { get; set; }

    public string? Message { get; set; }
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, PropertyNameCaseInsensitive = true)]
[JsonSerializable(typeof(GraphEventPage))]
[JsonSerializable(typeof(GraphErrorResponse))]
internal sealed partial class GraphJsonContext : JsonSerializerContext;
