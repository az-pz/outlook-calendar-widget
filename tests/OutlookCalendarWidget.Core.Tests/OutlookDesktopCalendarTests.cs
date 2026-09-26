using OutlookCalendarWidget.Core;
using static OutlookCalendarWidget.Core.Tests.TestData;

namespace OutlookCalendarWidget.Core.Tests;

public class OutlookDesktopCalendarTests
{
    private const string TeamsUrl = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0";

    private static OutlookAppointment Item(
        string subject,
        DateTimeOffset start,
        DateTimeOffset end,
        bool allDay = false,
        int busy = 2,
        int response = 3,
        int meeting = 3,
        string? location = null,
        string? teams = null,
        string entryId = "00000000AB12") => new()
        {
            EntryId = entryId,
            Subject = subject,
            StartUtc = start.UtcDateTime,
            EndUtc = end.UtcDateTime,
            AllDayEvent = allDay,
            BusyStatus = busy,
            ResponseStatus = response,
            MeetingStatus = meeting,
            Location = location,
            TeamsMeetingUrl = teams,
        };

    [Fact]
    public void Restrict_filter_uses_local_wall_clock_times_in_a_culture_neutral_format()
    {
        var (start, end) = AgendaBuilder.FetchWindow(Now, Eastern);

        var filter = OutlookDesktopCalendar.RestrictFilter(start, end, Eastern);

        Assert.Equal("[Start] < '2026-10-03 00:00' AND [End] > '2026-09-26 00:00'", filter);
    }

    [Fact]
    public void Restrict_filter_converts_other_offsets_to_the_calendar_time_zone()
    {
        var utcStart = new DateTimeOffset(2026, 9, 26, 4, 0, 0, TimeSpan.Zero);

        var filter = OutlookDesktopCalendar.RestrictFilter(utcStart, utcStart.AddHours(2), Eastern);

        Assert.Equal("[Start] < '2026-09-26 02:00' AND [End] > '2026-09-26 00:00'", filter);
    }

    [Fact]
    public void Maps_times_to_local_and_trims_text()
    {
        var e = OutlookDesktopCalendar.ToEvent(Item("  Standup ", At(26, 11, 35), At(26, 12), location: " Room 12 "), Eastern);

        Assert.Equal("Standup", e.Subject);
        Assert.Equal(At(26, 11, 35), e.Start);
        Assert.Equal(TimeSpan.FromHours(-4), e.Start.Offset);
        Assert.Equal(At(26, 12), e.End);
        Assert.Equal("Room 12", e.Location);
        Assert.False(e.IsAllDay);
        Assert.False(e.IsCancelled);
    }

    [Fact]
    public void All_day_events_start_at_local_midnight()
    {
        var e = OutlookDesktopCalendar.ToEvent(Item("Holiday", At(30, 0), At(30, 0).AddDays(1), allDay: true, busy: 0, response: 0, meeting: 0), Eastern);

        Assert.True(e.IsAllDay);
        Assert.Equal(new DateTime(2026, 9, 30), e.Start.DateTime);
        Assert.Equal(FreeBusyStatus.Free, e.ShowAs);
        Assert.Equal(MeetingResponse.None, e.Response);
    }

    [Theory]
    [InlineData(0, FreeBusyStatus.Free)]
    [InlineData(1, FreeBusyStatus.Tentative)]
    [InlineData(2, FreeBusyStatus.Busy)]
    [InlineData(3, FreeBusyStatus.Oof)]
    [InlineData(4, FreeBusyStatus.WorkingElsewhere)]
    [InlineData(99, FreeBusyStatus.Unknown)]
    public void Maps_busy_status(int value, FreeBusyStatus expected) =>
        Assert.Equal(expected, OutlookDesktopCalendar.MapBusyStatus(value));

    [Theory]
    [InlineData(0, MeetingResponse.None)]
    [InlineData(1, MeetingResponse.Organizer)]
    [InlineData(2, MeetingResponse.TentativelyAccepted)]
    [InlineData(3, MeetingResponse.Accepted)]
    [InlineData(4, MeetingResponse.Declined)]
    [InlineData(5, MeetingResponse.NotResponded)]
    public void Maps_response_status(int value, MeetingResponse expected) =>
        Assert.Equal(expected, OutlookDesktopCalendar.MapResponse(value));

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(3, false)]
    [InlineData(5, true)]
    [InlineData(7, true)]
    public void Detects_cancelled_meetings(int meetingStatus, bool cancelled) =>
        Assert.Equal(cancelled, OutlookDesktopCalendar.IsCancelled(meetingStatus));

    [Fact]
    public void Join_url_prefers_the_teams_property()
    {
        Assert.Equal(TeamsUrl, OutlookDesktopCalendar.FindJoinUrl(TeamsUrl, "Microsoft Teams Meeting"));
        Assert.Equal(TeamsUrl, OutlookDesktopCalendar.FindJoinUrl(TeamsUrl, "https://zoom.us/j/123"));
    }

    [Theory]
    [InlineData("https://zoom.us/j/123456789?pwd=abc", "https://zoom.us/j/123456789?pwd=abc")]
    [InlineData("Room 4; https://meet.google.com/abc-defg-hij", "https://meet.google.com/abc-defg-hij")]
    [InlineData("<https://contoso.webex.com/meet/pat>", "https://contoso.webex.com/meet/pat")]
    [InlineData("Microsoft Teams Meeting", null)]
    [InlineData("Building 7 / Room 1201", null)]
    [InlineData("ftp://files.contoso.com", null)]
    [InlineData("", null)]
    [InlineData(null, null)]
    public void Join_url_falls_back_to_a_link_in_the_location(string? location, string? expected) =>
        Assert.Equal(expected, OutlookDesktopCalendar.FindJoinUrl(null, location));

    [Fact]
    public void Web_link_opens_the_occurrence_in_outlook()
    {
        var e = OutlookDesktopCalendar.ToEvent(Item("Scrum", At(28, 12), At(28, 12, 30), entryId: "00AB/+="), Eastern);

        Assert.Equal("outlook-calendar-widget:open?id=00AB%2F%2B%3D&start=20260928T160000Z", e.WebLink);
        Assert.True(AppLinks.TryParseOpenOutlook(e.WebLink, out var request));
        Assert.Equal("00AB/+=", request.EntryId);
        Assert.Equal(At(28, 12), request.Start);
    }

    [Theory]
    [InlineData("outlook-calendar-widget:open")]
    [InlineData("outlook-calendar-widget:open/")]
    [InlineData("OUTLOOK-CALENDAR-WIDGET:OPEN?")]
    public void Plain_open_link_asks_for_the_calendar(string uri)
    {
        Assert.True(AppLinks.TryParseOpenOutlook(uri, out var request));
        Assert.Null(request.EntryId);
        Assert.Null(request.Start);
    }

    [Theory]
    [InlineData("outlook-calendar-widget:signin")]
    [InlineData("outlook-calendar-widget:settings")]
    [InlineData("outlook-calendar-widget:openx")]
    [InlineData("https://outlook.office.com/calendar")]
    [InlineData(null)]
    public void Other_links_are_not_open_requests(string? uri) =>
        Assert.False(AppLinks.TryParseOpenOutlook(uri, out _));

    [Fact]
    public void Open_link_ignores_a_malformed_start()
    {
        Assert.True(AppLinks.TryParseOpenOutlook("outlook-calendar-widget:open?id=ABC&start=tomorrow", out var request));
        Assert.Equal("ABC", request.EntryId);
        Assert.Null(request.Start);
    }

    [Fact]
    public void Occurrences_of_a_series_get_distinct_ids()
    {
        var first = OutlookDesktopCalendar.ToEvent(Item("Scrum", At(28, 12), At(28, 12, 30)), Eastern);
        var second = OutlookDesktopCalendar.ToEvent(Item("Scrum", At(29, 12), At(29, 12, 30)), Eastern);

        Assert.NotEqual(first.Id, second.Id);
        Assert.StartsWith("00000000AB12@", first.Id, StringComparison.Ordinal);
    }

    [Fact]
    public void To_events_filters_to_the_window_removes_duplicates_and_sorts()
    {
        var (start, end) = AgendaBuilder.FetchWindow(Now, Eastern);
        var items = new[]
        {
            Item("Later", At(26, 15), At(26, 16)),
            Item("Holiday", At(30, 0), At(30, 0).AddDays(1), allDay: true, entryId: "A"),
            Item("Holiday", At(30, 0), At(30, 0).AddDays(1), allDay: true, entryId: "B"),
            Item("Yesterday", At(25, 9), At(25, 10)),
            Item("Next week", At(26, 9).AddDays(8), At(26, 10).AddDays(8)),
            Item("Long running", At(16, 0), At(26, 0).AddDays(50)),
            Item("Earlier", At(26, 9), At(26, 9, 30)),
        };

        var events = OutlookDesktopCalendar.ToEvents(items, start, end, Eastern);

        Assert.Equal(["Long running", "Earlier", "Later", "Holiday"], events.Select(e => e.Subject));
    }

    [Fact]
    public void Real_outlook_items_flow_through_the_agenda()
    {
        var (start, end) = AgendaBuilder.FetchWindow(Now, Eastern);
        var events = OutlookDesktopCalendar.ToEvents(
            [
                Item("Canceled: Risk review", At(26, 11), At(26, 12), busy: 0, meeting: 7, teams: TeamsUrl),
                Item("Design sync", At(26, 10, 40), At(26, 11, 30), response: 5, busy: 1, location: "Microsoft Teams Meeting", teams: TeamsUrl),
                Item("Lunch", At(26, 12), At(26, 13), response: 0, meeting: 0),
            ],
            start,
            end,
            Eastern);

        var card = AgendaBuilder.Build(events, Now, WidgetPreferences.Default, WidgetSize.Medium, Eastern, EnUs, AppLinks.OpenOutlook);

        var rows = card.Days.Single().Events;
        Assert.Equal(["Design sync", "Lunch"], rows.Select(r => r.Subject));
        Assert.Equal("Design sync", card.Next!.Subject);
        Assert.True(card.Next.ShowJoin);
        Assert.Equal(TeamsUrl, card.Next.PrimaryUrl);
        Assert.Equal("Microsoft Teams Meeting · Not responded", rows[0].Detail);
        Assert.StartsWith(AppLinks.OpenOutlook + "?id=", rows[1].OpenUrl, StringComparison.Ordinal);
        Assert.Equal(AppLinks.OpenOutlook, card.OutlookUrl);
    }
}
