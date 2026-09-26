using System.Globalization;
using OutlookCalendarWidget.Core;

namespace OutlookCalendarWidget.Core.Tests;

internal static class TestData
{
    public static readonly TimeZoneInfo Eastern = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
    public static readonly CultureInfo EnUs = CultureInfo.GetCultureInfo("en-US");

    /// <summary>Saturday, September 26, 2026, 10:30 AM Eastern (UTC-4).</summary>
    public static readonly DateTimeOffset Now = new(2026, 9, 26, 10, 30, 0, TimeSpan.FromHours(-4));

    public static DateTimeOffset At(int day, int hour, int minute = 0) =>
        new(2026, 9, day, hour, minute, 0, TimeSpan.FromHours(-4));

    public static CalendarEvent Event(
        string subject,
        DateTimeOffset start,
        DateTimeOffset end,
        bool allDay = false,
        FreeBusyStatus showAs = FreeBusyStatus.Busy,
        MeetingResponse response = MeetingResponse.Accepted,
        bool cancelled = false,
        string? location = null,
        string? joinUrl = null,
        string? webLink = null) => new()
        {
            Id = subject,
            Subject = subject,
            Start = start,
            End = end,
            IsAllDay = allDay,
            ShowAs = showAs,
            Response = response,
            IsCancelled = cancelled,
            Location = location,
            JoinUrl = joinUrl,
            WebLink = webLink ?? $"https://outlook.office.com/calendar/item/{Uri.EscapeDataString(subject)}",
        };
}
