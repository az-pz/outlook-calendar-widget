namespace OutlookCalendarWidget.Core;

/// <summary>
/// A believable, always-current demo calendar. Lets people try the widget before connecting a Microsoft account, and
/// is used to render the widget picker screenshots.
/// </summary>
public static class SampleCalendar
{
    public const string AccountName = "Sample calendar";

    private const string TeamsLocation = "Microsoft Teams Meeting";
    private const string TeamsUrl = "https://teams.microsoft.com/";

    public static IReadOnlyList<CalendarEvent> Create(DateTimeOffset now, TimeZoneInfo timeZone)
    {
        ArgumentNullException.ThrowIfNull(timeZone);

        var local = TimeZoneInfo.ConvertTime(now, timeZone).DateTime;

        // The first meeting always starts within 15 minutes so the "Join" button is visible.
        var first = new DateTime(local.Year, local.Month, local.Day, local.Hour, local.Minute / 15 * 15, 0, DateTimeKind.Unspecified).AddMinutes(15);
        var tomorrow = local.Date.AddDays(1);

        return
        [
            Event(1, "Design review", At(first), At(first.AddMinutes(30)), TeamsLocation, joinUrl: TeamsUrl),
            Event(2, "1:1 with Priya", At(first.AddMinutes(60)), At(first.AddMinutes(90)), "Room 1120"),
            Event(3, "Lunch & learn: Windows widgets", At(first.AddMinutes(120)), At(first.AddMinutes(180)), "Commons", FreeBusyStatus.Tentative, MeetingResponse.TentativelyAccepted),
            Event(4, "Sprint planning", At(first.AddMinutes(210)), At(first.AddMinutes(270)), TeamsLocation, joinUrl: TeamsUrl),
            Event(5, "Team standup", At(tomorrow.AddHours(9.5)), At(tomorrow.AddHours(9.75)), TeamsLocation, joinUrl: TeamsUrl),
            Event(6, "Dentist", At(tomorrow.AddHours(13)), At(tomorrow.AddHours(14)), showAs: FreeBusyStatus.Oof),
            Event(7, "Hack week", At(tomorrow.AddDays(1)), At(tomorrow.AddDays(2)), showAs: FreeBusyStatus.Free, allDay: true),
        ];

        DateTimeOffset At(DateTime value) => new(value, timeZone.GetUtcOffset(value));
    }

    private static CalendarEvent Event(
        int id,
        string subject,
        DateTimeOffset start,
        DateTimeOffset end,
        string? location = null,
        FreeBusyStatus showAs = FreeBusyStatus.Busy,
        MeetingResponse response = MeetingResponse.Accepted,
        string? joinUrl = null,
        bool allDay = false) => new()
        {
            Id = $"sample-{id}",
            Subject = subject,
            Start = start,
            End = end,
            IsAllDay = allDay,
            Location = location,
            ShowAs = showAs,
            Response = response,
            JoinUrl = joinUrl,
        };
}
