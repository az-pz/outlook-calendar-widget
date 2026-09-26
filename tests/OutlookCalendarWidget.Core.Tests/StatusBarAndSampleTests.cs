using System.Buffers.Binary;
using OutlookCalendarWidget.Core;
using static OutlookCalendarWidget.Core.Tests.TestData;

namespace OutlookCalendarWidget.Core.Tests;

public class StatusBarAndSampleTests
{
    [Theory]
    [InlineData("accent")]
    [InlineData("warning")]
    [InlineData("attention")]
    [InlineData("good")]
    [InlineData("emphasis")]
    [InlineData("default")]
    public void Status_bar_is_a_valid_png_data_uri(string style)
    {
        const string prefix = "data:image/png;base64,";
        var uri = StatusBar.DataUri(style);

        Assert.StartsWith(prefix, uri, StringComparison.Ordinal);
        var png = Convert.FromBase64String(uri[prefix.Length..]);
        Assert.Equal([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A], png[..8]);
        Assert.Equal("IHDR", System.Text.Encoding.ASCII.GetString(png, 12, 4));
        Assert.Equal(8, BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(16)));
        Assert.Equal(72, BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(20)));
        Assert.Equal("IEND", System.Text.Encoding.ASCII.GetString(png, png.Length - 8, 4));
        Assert.True(uri.Length < 2048, "Keep widget payloads small.");
    }

    [Fact]
    public void Status_bar_uri_is_cached_and_differs_by_style()
    {
        Assert.Same(StatusBar.DataUri("accent"), StatusBar.DataUri("accent"));
        Assert.NotEqual(StatusBar.DataUri("accent"), StatusBar.DataUri("warning"));
        Assert.NotEqual(StatusBar.DataUri("accent"), StatusBar.DataUri("attention"));
    }

    [Fact]
    public void Sample_calendar_first_meeting_is_joinable()
    {
        var events = SampleCalendar.Create(Now, Eastern);

        var card = AgendaBuilder.Build(events, Now, WidgetPreferences.Default, WidgetSize.Medium, Eastern, EnUs);

        Assert.Equal("Design review", card.Next!.Subject);
        Assert.Equal("In 15 min", card.Next.RelativeText);
        Assert.True(card.Days.Single().Events[0].ShowJoin);
        Assert.Equal(3, card.Days.Single().Events.Count);
        Assert.Equal("+1 more", card.MoreText);
    }

    [Fact]
    public void Sample_calendar_spans_several_days_and_is_all_in_the_future()
    {
        var events = SampleCalendar.Create(Now, Eastern);

        Assert.All(events, e => Assert.True(e.Start > Now));
        Assert.Contains(events, e => e.IsAllDay);
        var week = AgendaBuilder.Build(events, Now, new WidgetPreferences { Range = AgendaRange.Week }, WidgetSize.Large, Eastern, EnUs, maxEvents: 20);
        Assert.Equal(3, week.Days.Count);
    }
}
