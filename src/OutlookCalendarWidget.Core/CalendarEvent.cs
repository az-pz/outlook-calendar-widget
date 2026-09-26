namespace OutlookCalendarWidget.Core;

/// <summary>Outlook's "show as" free/busy state for an event.</summary>
public enum FreeBusyStatus
{
    Unknown,
    Free,
    Tentative,
    Busy,
    Oof,
    WorkingElsewhere,
}

/// <summary>The signed-in user's response to a meeting.</summary>
public enum MeetingResponse
{
    None,
    Organizer,
    TentativelyAccepted,
    Accepted,
    Declined,
    NotResponded,
}

/// <summary>A calendar event normalized to the user's local time zone.</summary>
public sealed record CalendarEvent
{
    public required string Id { get; init; }

    public required string Subject { get; init; }

    public required DateTimeOffset Start { get; init; }

    public required DateTimeOffset End { get; init; }

    public bool IsAllDay { get; init; }

    public string? Location { get; init; }

    public FreeBusyStatus ShowAs { get; init; } = FreeBusyStatus.Busy;

    public MeetingResponse Response { get; init; } = MeetingResponse.None;

    public bool IsCancelled { get; init; }

    public string? JoinUrl { get; init; }

    public string? WebLink { get; init; }
}
