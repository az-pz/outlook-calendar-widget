using OutlookCalendarWidget.Core;
using static OutlookCalendarWidget.Core.Tests.TestData;

namespace OutlookCalendarWidget.Core.Tests;

public class AgendaBuilderTests
{
    private static CardData Build(
        IEnumerable<CalendarEvent> events,
        WidgetPreferences? preferences = null,
        WidgetSize size = WidgetSize.Large,
        DateTimeOffset? now = null) =>
        AgendaBuilder.Build(events, now ?? Now, preferences ?? WidgetPreferences.Default, size, Eastern, EnUs);

    [Fact]
    public void Filters_past_cancelled_declined_and_out_of_range_events()
    {
        var events = new[]
        {
            Event("Past", At(26, 8), At(26, 9)),
            Event("Cancelled", At(26, 11), At(26, 12), cancelled: true),
            Event("Declined", At(26, 12), At(26, 13), response: MeetingResponse.Declined),
            Event("Tomorrow", At(27, 9), At(27, 10)),
            Event("Upcoming", At(26, 14), At(26, 15)),
        };

        var card = Build(events);

        var subjects = card.Days.SelectMany(d => d.Events).Select(e => e.Subject).ToArray();
        Assert.Equal(["Upcoming"], subjects);
        Assert.Equal("1 event left", card.Summary);
    }

    [Fact]
    public void Preferences_can_include_declined_and_hide_all_day_events()
    {
        var events = new[]
        {
            Event("Holiday", At(26, 0), At(27, 0), allDay: true, showAs: FreeBusyStatus.Free),
            Event("Declined", At(26, 12), At(26, 13), response: MeetingResponse.Declined),
        };

        var card = Build(events, new WidgetPreferences { ShowDeclined = true, ShowAllDay = false });

        var item = Assert.Single(card.Days.SelectMany(d => d.Events));
        Assert.Equal("Declined", item.Subject);
        Assert.Equal("Declined", item.StatusText);
        Assert.Equal("default", item.StatusStyle);
    }

    [Fact]
    public void Orders_all_day_first_then_by_start_and_marks_ongoing_event()
    {
        var events = new[]
        {
            Event("Later", At(26, 15), At(26, 16)),
            Event("Standup", At(26, 10), At(26, 11), joinUrl: "https://teams.microsoft.com/l/meetup-join/1"),
            Event("Holiday", At(26, 0), At(27, 0), allDay: true),
        };

        var card = Build(events);

        var items = card.Days.Single().Events;
        Assert.Equal(["Holiday", "Standup", "Later"], items.Select(e => e.Subject));
        Assert.Equal("All day", items[0].TimeText);
        Assert.False(items[0].HasEndText);

        var standup = items[1];
        Assert.True(standup.IsNow);
        Assert.True(standup.ShowJoin);
        Assert.Equal("emphasis", standup.RowStyle);
        Assert.Equal("https://teams.microsoft.com/l/meetup-join/1", standup.PrimaryUrl);
        Assert.Equal("Now · ends in 30 min", standup.RelativeText);
        Assert.Equal(At(26, 10).ToString("t", EnUs), standup.TimeText);
        Assert.Equal(At(26, 11).ToString("t", EnUs), standup.EndText);

        Assert.False(items[2].ShowJoin);
        Assert.Equal("default", items[2].RowStyle);
    }

    [Fact]
    public void Next_event_skips_all_day_free_and_declined_items()
    {
        var events = new[]
        {
            Event("Holiday", At(26, 0), At(27, 0), allDay: true),
            Event("Focus time", At(26, 11), At(26, 12), showAs: FreeBusyStatus.Free),
            Event("Declined", At(26, 11), At(26, 12), response: MeetingResponse.Declined),
            Event("Design review", At(26, 13), At(26, 14), location: "Room 42"),
        };

        var card = Build(events, new WidgetPreferences { ShowDeclined = true });

        Assert.True(card.HasNext);
        Assert.Equal("Design review", card.Next!.Subject);
        Assert.Equal("In 2 hr 30 min", card.Next.RelativeText);
        Assert.Equal("Room 42", card.Next.Detail);
    }

    [Fact]
    public void Next_event_falls_back_to_free_events_when_nothing_else_is_scheduled()
    {
        var card = Build([Event("Focus time", At(26, 11), At(26, 12), showAs: FreeBusyStatus.Free)]);

        Assert.Equal("Focus time", card.Next!.Subject);
        Assert.Equal("emphasis", card.Next.StatusStyle);
    }

    [Theory]
    [InlineData(WidgetSize.Small, 0, "")]
    [InlineData(WidgetSize.Medium, 3, "+5 more")]
    [InlineData(WidgetSize.Large, 6, "+2 more")]
    public void Caps_the_number_of_events_by_widget_size(WidgetSize size, int expectedShown, string expectedMore)
    {
        var events = Enumerable.Range(0, 8)
            .Select(i => Event($"Meeting {i}", At(26, 11 + i), At(26, 11 + i, 30)))
            .ToArray();

        var card = Build(events, size: size);

        Assert.Equal(expectedShown, card.Days.SelectMany(d => d.Events).Count());
        Assert.Equal(expectedMore, card.MoreText);
        Assert.Equal(expectedMore.Length > 0, card.HasMore);
        Assert.Equal("Meeting 0", card.Next!.Subject);
        Assert.Equal("8 events left", card.Summary);
    }

    [Fact]
    public void Groups_by_day_for_multi_day_ranges()
    {
        var events = new[]
        {
            Event("Today", At(26, 13), At(26, 14)),
            Event("Tomorrow", At(27, 9), At(27, 10)),
            Event("Monday", At(28, 9), At(28, 10)),
        };

        var twoDays = Build(events, new WidgetPreferences { Range = AgendaRange.TwoDays });
        Assert.True(twoDays.ShowDayHeaders);
        Assert.Equal("Today & tomorrow", twoDays.Heading);
        Assert.Equal(["Today", "Tomorrow"], twoDays.Days.Select(d => d.Label));
        Assert.Equal("2 upcoming", twoDays.Summary);

        var week = Build(events, new WidgetPreferences { Range = AgendaRange.Week });
        Assert.Equal(["Today", "Tomorrow", "Monday, September 28"], week.Days.Select(d => d.Label));
        Assert.Equal("Next 7 days", week.Heading);
    }

    [Fact]
    public void Multi_day_event_that_started_earlier_shows_under_today_with_end_time()
    {
        var card = Build([Event("Offsite", At(25, 9), At(26, 17))]);

        var item = card.Days.Single().Events.Single();
        Assert.Equal("Today", card.Days.Single().Label);
        Assert.Equal("Until", item.TimeText);
        Assert.Equal(At(26, 17).ToString("t", EnUs), item.EndText);
    }

    [Fact]
    public void Timed_event_spanning_the_whole_day_is_shown_as_all_day_and_never_next()
    {
        var events = new[]
        {
            Event("Parental leave", new DateTimeOffset(2026, 9, 16, 9, 0, 0, TimeSpan.FromHours(-4)), new DateTimeOffset(2026, 11, 17, 17, 0, 0, TimeSpan.FromHours(-5)), showAs: FreeBusyStatus.Oof),
            Event("Late call", At(26, 22), At(27, 1)),
        };

        var card = Build(events);

        var items = card.Days.Single().Events;
        Assert.Equal(["Parental leave", "Late call"], items.Select(e => e.Subject));
        Assert.Equal("All day", items[0].TimeText);
        Assert.Equal("All day", items[0].RelativeText);
        Assert.False(items[0].HasEndText);
        Assert.False(items[0].IsNow);
        Assert.True(items[0].IsAllDay);
        Assert.Equal("Late call", card.Next!.Subject);
        Assert.Equal(At(27, 1).ToString("t", EnUs), items[1].EndText);

        var hidden = Build(events, new WidgetPreferences { ShowAllDay = false });
        Assert.Equal(["Late call"], hidden.Days.SelectMany(d => d.Events).Select(e => e.Subject));
    }

    [Fact]
    public void Empty_calendar_produces_friendly_empty_state()
    {
        var card = Build([]);

        Assert.True(card.IsOk);
        Assert.False(card.HasEvents);
        Assert.False(card.HasNext);
        Assert.Equal("You're all clear", card.EmptyTitle);
        Assert.Equal("No more meetings today.", card.EmptyText);
        Assert.Equal("No events left", card.Summary);
        Assert.Equal("Saturday, September 26", card.DateLabel);
        Assert.Equal("Today", card.Heading);
    }

    [Fact]
    public void Uses_outlook_calendar_url_when_event_has_no_web_link()
    {
        var evt = Event("No link", At(26, 13), At(26, 14)) with { WebLink = null };

        var card = AgendaBuilder.Build([evt], Now, WidgetPreferences.Default, WidgetSize.Medium, Eastern, EnUs, "https://example.test/cal");

        Assert.Equal("https://example.test/cal", card.Next!.OpenUrl);
        Assert.Equal("https://example.test/cal", card.OutlookUrl);
    }

    [Theory]
    [InlineData(26, 10, 31, "Starting now")]
    [InlineData(26, 10, 55, "In 25 min")]
    [InlineData(26, 12, 30, "In 2 hr")]
    [InlineData(26, 13, 15, "In 2 hr 45 min")]
    public void Relative_text_for_upcoming_events_today(int day, int hour, int minute, string expected)
    {
        var start = At(day, hour, minute);
        var evt = Event("x", start, start.AddMinutes(30));

        Assert.Equal(expected, AgendaBuilder.RelativeText(evt, Now.AddSeconds(30), AgendaBuilder.StartOfDay(Now, Eastern), Eastern, EnUs));
    }

    [Fact]
    public void Relative_text_for_later_and_future_days()
    {
        var today = AgendaBuilder.StartOfDay(Now, Eastern);

        var evening = Event("Dinner", At(26, 19), At(26, 20));
        Assert.Equal($"At {At(26, 19).ToString("t", EnUs)}", AgendaBuilder.RelativeText(evening, Now, today, Eastern, EnUs));

        var tomorrow = Event("Brunch", At(27, 9), At(27, 10));
        Assert.Equal($"Tomorrow at {At(27, 9).ToString("t", EnUs)}", AgendaBuilder.RelativeText(tomorrow, Now, today, Eastern, EnUs));

        var monday = Event("Planning", At(28, 9), At(28, 10));
        Assert.Equal($"Mon at {At(28, 9).ToString("t", EnUs)}", AgendaBuilder.RelativeText(monday, Now, today, Eastern, EnUs));

        var longMeeting = Event("Workshop", At(26, 9), At(26, 13));
        Assert.Equal($"Now · until {At(26, 13).ToString("t", EnUs)}", AgendaBuilder.RelativeText(longMeeting, Now, today, Eastern, EnUs));
    }

    [Fact]
    public void Join_button_appears_only_within_fifteen_minutes_of_start()
    {
        var soon = Event("Soon", At(26, 10, 44), At(26, 11), joinUrl: "https://teams.microsoft.com/1");
        var later = Event("Later", At(26, 10, 46), At(26, 11), joinUrl: "https://teams.microsoft.com/2");

        var card = Build([soon, later]);

        var items = card.Days.Single().Events;
        Assert.True(items[0].ShowJoin);
        Assert.False(items[1].ShowJoin);
        Assert.Equal(items[1].OpenUrl, items[1].PrimaryUrl);
    }

    [Fact]
    public void Fetch_window_starts_at_local_midnight_and_spans_a_week()
    {
        var (start, end) = AgendaBuilder.FetchWindow(Now, Eastern);

        Assert.Equal(new DateTimeOffset(2026, 9, 26, 0, 0, 0, TimeSpan.FromHours(-4)), start);
        Assert.Equal(new DateTimeOffset(2026, 10, 3, 0, 0, 0, TimeSpan.FromHours(-4)), end);
    }

    [Fact]
    public void Day_math_handles_daylight_saving_transitions()
    {
        // US DST ends on Sunday, November 1, 2026.
        var saturday = new DateTimeOffset(2026, 10, 31, 22, 0, 0, TimeSpan.FromHours(-4));
        var sundayMeeting = Event("Sunday", new DateTimeOffset(2026, 11, 1, 9, 0, 0, TimeSpan.FromHours(-5)), new DateTimeOffset(2026, 11, 1, 10, 0, 0, TimeSpan.FromHours(-5)));

        var card = AgendaBuilder.Build([sundayMeeting], saturday, new WidgetPreferences { Range = AgendaRange.TwoDays }, WidgetSize.Large, Eastern, EnUs);

        Assert.Equal("Tomorrow", card.Days.Single().Label);
        var (_, end) = AgendaBuilder.FetchWindow(saturday, Eastern);
        Assert.Equal(TimeSpan.FromHours(-5), end.Offset);
    }
}
