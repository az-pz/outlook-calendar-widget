using System.Globalization;

namespace OutlookCalendarWidget.Core;

/// <summary>Raw values read from a classic Outlook <c>AppointmentItem</c> through the Outlook object model.</summary>
public sealed record OutlookAppointment
{
    public required string EntryId { get; init; }

    public string? Subject { get; init; }

    public required DateTime StartUtc { get; init; }

    public required DateTime EndUtc { get; init; }

    public bool AllDayEvent { get; init; }

    /// <summary><c>OlBusyStatus</c>: 0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere.</summary>
    public int BusyStatus { get; init; } = 2;

    /// <summary><c>OlResponseStatus</c>: 0 none, 1 organizer, 2 tentative, 3 accepted, 4 declined, 5 not responded.</summary>
    public int ResponseStatus { get; init; }

    /// <summary><c>OlMeetingStatus</c>: 0 appointment, 1 meeting, 3 received, 5 canceled, 7 received and canceled.</summary>
    public int MeetingStatus { get; init; }

    public bool IsRecurring { get; init; }

    public string? Location { get; init; }

    /// <summary>Value of the <see cref="OutlookDesktopCalendar.TeamsMeetingUrlProperty"/> named property, if set.</summary>
    public string? TeamsMeetingUrl { get; init; }
}

/// <summary>
/// Maps classic Outlook calendar items to <see cref="CalendarEvent"/>. The COM calls themselves live in the app
/// (they need a running Outlook); everything here is pure so it can be unit tested.
/// </summary>
public static class OutlookDesktopCalendar
{
    /// <summary><c>OlDefaultFolders.olFolderCalendar</c>.</summary>
    public const int CalendarFolderId = 9;

    /// <summary><c>OlObjectClass.olAppointment</c>.</summary>
    public const int AppointmentClass = 26;

    /// <summary>MAPI named property (PS_PUBLIC_STRINGS) where Outlook stores a Teams meeting's join link.</summary>
    public const string TeamsMeetingUrlProperty =
        "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/SkypeTeamsMeetingUrl";

    /// <summary>
    /// A Jet filter for <c>Items.Restrict</c> matching items that overlap the window. Outlook compares these values in
    /// local time; the ISO-like format parses the same way regardless of the user's regional date format.
    /// </summary>
    public static string RestrictFilter(DateTimeOffset start, DateTimeOffset end, TimeZoneInfo timeZone)
    {
        ArgumentNullException.ThrowIfNull(timeZone);
        return $"[Start] < '{FormatLocal(end, timeZone)}' AND [End] > '{FormatLocal(start, timeZone)}'";
    }

    /// <summary>Maps, de-duplicates and sorts the items, keeping only those that overlap the window.</summary>
    public static IReadOnlyList<CalendarEvent> ToEvents(
        IEnumerable<OutlookAppointment> items,
        DateTimeOffset start,
        DateTimeOffset end,
        TimeZoneInfo timeZone)
    {
        ArgumentNullException.ThrowIfNull(items);
        ArgumentNullException.ThrowIfNull(timeZone);

        return items
            .Select(item => ToEvent(item, timeZone))
            .Where(e => e.Start < end && e.End > start)

            // Subscribed holiday calendars often add the same all-day event twice.
            .DistinctBy(e => (e.Subject, e.Start, e.End, e.IsAllDay))
            .OrderBy(e => e.Start)
            .ThenBy(e => e.End)
            .ToList();
    }

    public static CalendarEvent ToEvent(OutlookAppointment item, TimeZoneInfo timeZone)
    {
        ArgumentNullException.ThrowIfNull(item);
        ArgumentNullException.ThrowIfNull(timeZone);

        var start = ToLocal(item.StartUtc, timeZone);
        var end = ToLocal(item.EndUtc, timeZone);
        return new CalendarEvent
        {
            // Occurrences of a recurring series share the master's EntryID.
            Id = $"{item.EntryId}@{start.UtcDateTime.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture)}",
            Subject = string.IsNullOrWhiteSpace(item.Subject) ? string.Empty : item.Subject.Trim(),
            Start = start,
            End = end < start ? start : end,
            IsAllDay = item.AllDayEvent,
            IsCancelled = IsCancelled(item.MeetingStatus),
            Location = string.IsNullOrWhiteSpace(item.Location) ? null : item.Location.Trim(),
            ShowAs = MapBusyStatus(item.BusyStatus),
            Response = MapResponse(item.ResponseStatus),
            JoinUrl = FindJoinUrl(item.TeamsMeetingUrl, item.Location),
            WebLink = string.IsNullOrEmpty(item.EntryId) ? null : AppLinks.OpenOutlookItem(item.EntryId, start),
        };
    }

    internal static FreeBusyStatus MapBusyStatus(int value) => value switch
    {
        0 => FreeBusyStatus.Free,
        1 => FreeBusyStatus.Tentative,
        2 => FreeBusyStatus.Busy,
        3 => FreeBusyStatus.Oof,
        4 => FreeBusyStatus.WorkingElsewhere,
        _ => FreeBusyStatus.Unknown,
    };

    internal static MeetingResponse MapResponse(int value) => value switch
    {
        1 => MeetingResponse.Organizer,
        2 => MeetingResponse.TentativelyAccepted,
        3 => MeetingResponse.Accepted,
        4 => MeetingResponse.Declined,
        5 => MeetingResponse.NotResponded,
        _ => MeetingResponse.None,
    };

    internal static bool IsCancelled(int meetingStatus) => meetingStatus is 5 or 7;

    /// <summary>The Teams join link, or else a meeting link typed into the location (Zoom, Webex, Meet, …).</summary>
    internal static string? FindJoinUrl(string? teamsMeetingUrl, string? location)
    {
        if (IsWebUrl(teamsMeetingUrl?.Trim(), out var teams))
        {
            return teams;
        }

        if (string.IsNullOrWhiteSpace(location))
        {
            return null;
        }

        foreach (var token in location.Split([' ', ';', ',', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
        {
            if (IsWebUrl(token.Trim('<', '>', '(', ')', '"'), out var url))
            {
                return url;
            }
        }

        return null;
    }

    private static bool IsWebUrl(string? value, out string url)
    {
        url = string.Empty;
        if (string.IsNullOrEmpty(value) ||
            !Uri.TryCreate(value, UriKind.Absolute, out var uri) ||
            (uri.Scheme != Uri.UriSchemeHttps && uri.Scheme != Uri.UriSchemeHttp))
        {
            return false;
        }

        url = value;
        return true;
    }

    private static DateTimeOffset ToLocal(DateTime utc, TimeZoneInfo timeZone) =>
        TimeZoneInfo.ConvertTime(new DateTimeOffset(DateTime.SpecifyKind(utc, DateTimeKind.Utc)), timeZone);

    private static string FormatLocal(DateTimeOffset value, TimeZoneInfo timeZone) =>
        TimeZoneInfo.ConvertTime(value, timeZone).ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
}
