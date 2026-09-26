#:project ../../src/OutlookCalendarWidget.Core/OutlookCalendarWidget.Core.csproj

// Prints the widget card data for the sample calendar at a fixed moment, so picker screenshots are reproducible.
// Usage: dotnet run SampleCard.cs -- [small|medium|large]
using System.Globalization;
using OutlookCalendarWidget.Core;

var size = args.Length > 0 ? Enum.Parse<WidgetSize>(args[0], ignoreCase: true) : WidgetSize.Medium;
var timeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
var now = new DateTimeOffset(2026, 9, 29, 9, 50, 0, TimeSpan.FromHours(-4));

var card = AgendaBuilder.Build(
    SampleCalendar.Create(now, timeZone),
    now,
    WidgetPreferences.Default,
    size,
    timeZone,
    CultureInfo.GetCultureInfo("en-US"));

Console.WriteLine(card.ToJson());
