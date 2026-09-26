using System.Globalization;

namespace OutlookCalendarWidget.Core;

/// <summary>Turns raw calendar events into the size-aware <see cref="CardData"/> rendered by the widget.</summary>
public static class AgendaBuilder
{
    private static readonly TimeSpan JoinWindow = TimeSpan.FromMinutes(15);

    public static int MaxEventsFor(WidgetSize size) => size switch
    {
        WidgetSize.Small => 0,
        WidgetSize.Medium => 3,
        _ => 6,
    };

    /// <summary>The fetch window needed to satisfy any widget's preferences, starting at local midnight today.</summary>
    public static (DateTimeOffset Start, DateTimeOffset End) FetchWindow(DateTimeOffset now, TimeZoneInfo timeZone)
    {
        var start = StartOfDay(now, timeZone);
        return (start, AddDays(start, 7, timeZone));
    }

    public static CardData Build(
        IEnumerable<CalendarEvent> events,
        DateTimeOffset now,
        WidgetPreferences preferences,
        WidgetSize size,
        TimeZoneInfo timeZone,
        CultureInfo culture,
        string outlookUrl = AppLinks.DefaultOutlookCalendarUrl,
        DateTimeOffset? lastUpdated = null,
        int? maxEvents = null)
    {
        ArgumentNullException.ThrowIfNull(events);
        ArgumentNullException.ThrowIfNull(preferences);
        ArgumentNullException.ThrowIfNull(timeZone);
        ArgumentNullException.ThrowIfNull(culture);

        now = TimeZoneInfo.ConvertTime(now, timeZone);
        var today = StartOfDay(now, timeZone);
        var rangeEnd = AddDays(today, preferences.Range switch
        {
            AgendaRange.TwoDays => 2,
            AgendaRange.Week => 7,
            _ => 1,
        }, timeZone);

        var visible = events
            .Where(e => !e.IsCancelled)
            .Where(e => e.End > now && e.Start < rangeEnd)
            .Where(e => preferences.ShowDeclined || e.Response != MeetingResponse.Declined)
            .Select(e => Place(e, today, timeZone))
            .Where(x => preferences.ShowAllDay || !x.WholeDay)
            .OrderBy(x => x.Day)
            .ThenByDescending(x => x.WholeDay)
            .ThenBy(x => x.Event.Start)
            .ThenBy(x => x.Event.End)
            .ToList();

        var next = visible
            .Where(x => !x.WholeDay && x.Event.Response != MeetingResponse.Declined && x.Event.ShowAs != FreeBusyStatus.Free)
            .FirstOrDefault()
            ?? visible.FirstOrDefault(x => !x.WholeDay);

        var shown = visible.Take(maxEvents ?? MaxEventsFor(size)).ToList();
        var hidden = visible.Count - shown.Count;

        var days = shown
            .GroupBy(x => x.Day)
            .Select(g => new CardDay
            {
                Label = DayLabel(g.Key, today, timeZone, culture),
                Events = g.Select(x => ToCardEvent(x, now, today, timeZone, culture, outlookUrl)).ToList(),
            })
            .ToList();

        return new CardData
        {
            State = CardStates.Ok,
            Heading = preferences.Range switch
            {
                AgendaRange.TwoDays => "Today & tomorrow",
                AgendaRange.Week => "Next 7 days",
                _ => "Today",
            },
            DateLabel = now.ToString("dddd, " + culture.DateTimeFormat.MonthDayPattern, culture),
            Summary = Summary(visible.Count, preferences.Range),
            Next = next is null ? null : ToCardEvent(next, now, today, timeZone, culture, outlookUrl),
            ShowDayHeaders = preferences.Range != AgendaRange.Today,
            Days = days,
            EmptyTitle = "You're all clear",
            EmptyText = preferences.Range == AgendaRange.Today ? "No more meetings today." : "Nothing on your calendar.",
            MoreText = hidden > 0 && size != WidgetSize.Small ? $"+{hidden} more" : string.Empty,
            UpdatedText = $"Updated {(lastUpdated ?? now).ToString("t", culture)}",
            OutlookUrl = outlookUrl,
        };
    }

    internal static string RelativeText(CalendarEvent e, DateTimeOffset now, DateTimeOffset today, TimeZoneInfo timeZone, CultureInfo culture)
    {
        if (e.IsAllDay)
        {
            return "All day";
        }

        if (e.Start <= now)
        {
            var remaining = e.End - now;
            return remaining.TotalMinutes < 60
                ? $"Now · ends in {Math.Max(1, (int)Math.Ceiling(remaining.TotalMinutes))} min"
                : $"Now · until {e.End.ToString("t", culture)}";
        }

        var until = e.Start - now;
        if (until.TotalMinutes < 1)
        {
            return "Starting now";
        }

        var startDay = StartOfDay(e.Start, timeZone);
        if (startDay == today)
        {
            if (until.TotalMinutes < 60)
            {
                return $"In {(int)Math.Ceiling(until.TotalMinutes)} min";
            }

            if (until.TotalHours < 6)
            {
                var totalMinutes = (int)Math.Ceiling(until.TotalMinutes);
                var hours = totalMinutes / 60;
                var minutes = totalMinutes % 60;
                return minutes == 0 ? $"In {hours} hr" : $"In {hours} hr {minutes} min";
            }

            return $"At {e.Start.ToString("t", culture)}";
        }

        if (startDay == AddDays(today, 1, timeZone))
        {
            return $"Tomorrow at {e.Start.ToString("t", culture)}";
        }

        return $"{e.Start.ToString("ddd", culture)} at {e.Start.ToString("t", culture)}";
    }

    private static CardEvent ToCardEvent(
        Placement placement,
        DateTimeOffset now,
        DateTimeOffset today,
        TimeZoneInfo timeZone,
        CultureInfo culture,
        string outlookUrl)
    {
        var (e, day, wholeDay) = placement;
        string timeText;
        var endText = string.Empty;
        if (wholeDay)
        {
            timeText = "All day";
        }
        else if (e.Start < day)
        {
            timeText = "Until";
            endText = e.End.ToString("t", culture);
        }
        else
        {
            timeText = e.Start.ToString("t", culture);
            endText = e.End.ToString("t", culture);
        }

        var statusText = e.Response switch
        {
            MeetingResponse.Declined => "Declined",
            MeetingResponse.NotResponded => "Not responded",
            MeetingResponse.TentativelyAccepted => "Tentative",
            _ => e.ShowAs == FreeBusyStatus.Tentative ? "Tentative" : string.Empty,
        };

        var location = e.Location ?? string.Empty;
        var detail = string.Join(" · ", new[] { location, statusText }.Where(s => s.Length > 0));

        return new CardEvent
        {
            Subject = string.IsNullOrWhiteSpace(e.Subject) ? "(No title)" : e.Subject,
            TimeText = timeText,
            EndText = endText,
            RelativeText = wholeDay ? "All day" : RelativeText(e, now, today, timeZone, culture),
            Location = location,
            StatusStyle = StatusStyle(e),
            StatusText = statusText,
            Detail = detail,
            OpenUrl = e.WebLink ?? outlookUrl,
            JoinUrl = e.JoinUrl ?? string.Empty,
            ShowJoin = e.JoinUrl is not null && !wholeDay && e.Start - now <= JoinWindow && e.End > now,
            IsNow = !wholeDay && e.Start <= now && e.End > now,
            IsAllDay = wholeDay,
        };
    }

    private static string StatusStyle(CalendarEvent e)
    {
        if (e.Response == MeetingResponse.Declined)
        {
            return "default";
        }

        return e.ShowAs switch
        {
            FreeBusyStatus.Tentative => "warning",
            FreeBusyStatus.Oof => "attention",
            FreeBusyStatus.WorkingElsewhere => "good",
            FreeBusyStatus.Free => "emphasis",
            _ => "accent",
        };
    }

    private static string Summary(int count, AgendaRange range)
    {
        if (range == AgendaRange.Today)
        {
            return count switch
            {
                0 => "No events left",
                1 => "1 event left",
                _ => $"{count} events left",
            };
        }

        return $"{count} upcoming";
    }

    private static string DayLabel(DateTimeOffset day, DateTimeOffset today, TimeZoneInfo timeZone, CultureInfo culture)
    {
        if (day == today)
        {
            return "Today";
        }

        if (day == AddDays(today, 1, timeZone))
        {
            return "Tomorrow";
        }

        return day.ToString("dddd, " + culture.DateTimeFormat.MonthDayPattern, culture);
    }

    /// <summary>
    /// Places an event on the first visible day it touches. Timed events that span that entire day
    /// (e.g. a multi-week "Out of office" block) are treated like all-day events.
    /// </summary>
    private static Placement Place(CalendarEvent e, DateTimeOffset today, TimeZoneInfo timeZone)
    {
        var start = StartOfDay(e.Start, timeZone);
        var day = start < today ? today : start;
        var wholeDay = e.IsAllDay || (e.Start <= day && e.End >= AddDays(day, 1, timeZone));
        return new Placement(e, day, wholeDay);
    }

    private sealed record Placement(CalendarEvent Event, DateTimeOffset Day, bool WholeDay);

    internal static DateTimeOffset StartOfDay(DateTimeOffset value, TimeZoneInfo timeZone)
    {
        var local = TimeZoneInfo.ConvertTime(value, timeZone);
        var midnight = local.Date;
        return new DateTimeOffset(midnight, timeZone.GetUtcOffset(midnight));
    }

    private static DateTimeOffset AddDays(DateTimeOffset startOfDay, int days, TimeZoneInfo timeZone)
    {
        var midnight = startOfDay.Date.AddDays(days);
        return new DateTimeOffset(midnight, timeZone.GetUtcOffset(midnight));
    }
}
