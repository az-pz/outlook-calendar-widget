using System.Text.Json;
using System.Text.Json.Nodes;
using AdaptiveCards.Templating;
using OutlookCalendarWidget.Core;
using static OutlookCalendarWidget.Core.Tests.TestData;

namespace OutlookCalendarWidget.Core.Tests;

/// <summary>
/// Expands the widget templates with the same Adaptive Card templating language the Widgets Board uses,
/// so binding mistakes (typos in field names, bad <c>$when</c> expressions) fail in CI instead of on the board.
/// </summary>
public class WidgetTemplateTests
{
    private static readonly CalendarEvent[] SampleEvents =
    [
        Event("Company holiday", At(26, 0), At(27, 0), allDay: true, showAs: FreeBusyStatus.Oof),
        Event("Standup", At(26, 10, 40), At(26, 11), location: "Teams", joinUrl: "https://teams.microsoft.com/l/meetup-join/1"),
        Event("Design review", At(26, 13), At(26, 14), location: "Room 42", response: MeetingResponse.TentativelyAccepted, showAs: FreeBusyStatus.Tentative),
        Event("1:1", At(26, 15), At(26, 15, 30)),
        Event("Tomorrow sync", At(27, 9), At(27, 10)),
    ];

    public static TheoryData<string> Sizes => ["small", "medium", "large"];

    [Theory]
    [MemberData(nameof(Sizes))]
    public void Agenda_binds_for_every_size(string size)
    {
        var data = AgendaBuilder.Build(SampleEvents, Now, WidgetPreferences.Default, Parse(size), Eastern, EnUs);

        var card = Expand(WidgetTemplates.Calendar, data.ToJson(), size);
        var texts = Texts(card);

        if (size == "small")
        {
            Assert.Contains("Standup", texts);
            Assert.Contains("In 10 min", texts);
            Assert.DoesNotContain("Design review", texts);
            Assert.Equal("https://teams.microsoft.com/l/meetup-join/1", SelectActionUrls(card).Single());
        }
        else
        {
            Assert.Contains("Today", texts);
            Assert.Contains("Saturday, September 26", texts);
            Assert.Contains("Company holiday", texts);
            Assert.Contains("Standup", texts);
            Assert.Contains("Teams", texts);
            Assert.Contains("Design review", texts);
            Assert.Contains("Room 42 · Tentative", texts);
            Assert.Contains("Join", ActionTitles(card));
            var bars = Descendants(card).Where(n => n["type"]?.GetValue<string>() == "Image").ToList();
            Assert.NotEmpty(bars);
            Assert.All(bars, bar => Assert.StartsWith("data:image/png;base64,", bar["url"]!.GetValue<string>(), StringComparison.Ordinal));
        }

        if (size == "medium")
        {
            Assert.Contains("+1 more", texts);
            Assert.DoesNotContain("1:1", texts);
        }

        if (size == "large")
        {
            Assert.Contains("1:1", texts);
            Assert.Contains("4 events left", texts);
            Assert.Contains("Open Outlook", ActionTitles(card));
            Assert.Contains(texts, t => t.StartsWith("Updated ", StringComparison.Ordinal));
        }
    }

    [Theory]
    [MemberData(nameof(Sizes))]
    public void Multi_day_agenda_shows_day_headers(string size)
    {
        var data = AgendaBuilder.Build(SampleEvents, Now, new WidgetPreferences { Range = AgendaRange.TwoDays }, Parse(size), Eastern, EnUs);

        var texts = Texts(Expand(WidgetTemplates.Calendar, data.ToJson(), size));

        if (size == "large")
        {
            Assert.Contains("Today & tomorrow", texts);
            Assert.Contains("Tomorrow", texts);
            Assert.Contains("Tomorrow sync", texts);
        }
        else if (size == "medium")
        {
            Assert.Contains("Today", texts);
            Assert.DoesNotContain("Tomorrow", texts);
        }
    }

    [Theory]
    [MemberData(nameof(Sizes))]
    public void Empty_agenda_binds(string size)
    {
        var data = AgendaBuilder.Build([], Now, WidgetPreferences.Default, Parse(size), Eastern, EnUs);

        var texts = Texts(Expand(WidgetTemplates.Calendar, data.ToJson(), size));

        Assert.Contains("You're all clear", texts);
        Assert.Contains("No more meetings today.", texts);
    }

    [Theory]
    [MemberData(nameof(Sizes))]
    public void Sign_in_state_uses_protocol_link(string size)
    {
        var data = CardData.ForStatus(CardStates.SignedOut, "Sign in to see your meetings", "Connect your Microsoft account.", "Sign in", actionUrl: AppLinks.SignIn);

        var card = Expand(WidgetTemplates.Calendar, data.ToJson(), size);
        var texts = Texts(card);

        Assert.Contains("Sign in to see your meetings", texts);
        Assert.Equal(size != "small", texts.Contains("Connect your Microsoft account."));
        var action = Actions(card).Single();
        Assert.Equal("Action.OpenUrl", action["type"]!.GetValue<string>());
        Assert.Equal("outlook-calendar-widget:signin", action["url"]!.GetValue<string>());
        Assert.DoesNotContain("Today", texts);
    }

    [Fact]
    public void Error_state_uses_execute_verb()
    {
        var data = CardData.ForStatus(CardStates.Error, "Couldn't load your calendar", "Network unavailable", "Retry", actionVerb: "refresh");

        var card = Expand(WidgetTemplates.Calendar, data.ToJson(), "medium");

        var action = Actions(card).Single();
        Assert.Equal("Action.Execute", action["type"]!.GetValue<string>());
        Assert.Equal("refresh", action["verb"]!.GetValue<string>());
        Assert.Equal("Retry", action["title"]!.GetValue<string>());
    }

    [Fact]
    public void Customization_template_binds_current_preferences()
    {
        var data = new WidgetPreferences { Range = AgendaRange.Week, ShowDeclined = true }.ToCustomizationData();

        var card = Expand(WidgetTemplates.Customize, data, "medium");

        var inputs = Descendants(card).Where(n => n["id"] is not null).ToDictionary(n => n["id"]!.GetValue<string>(), n => n["value"]!.GetValue<string>());
        Assert.Equal("week", inputs["range"]);
        Assert.Equal("true", inputs["showDeclined"]);
        Assert.Equal("true", inputs["showAllDay"]);
        Assert.Equal(["saveCustomization", "exitCustomization"], card["actions"]!.AsArray().Select(a => a!["verb"]!.GetValue<string>()));
    }

    private static WidgetSize Parse(string size) => Enum.Parse<WidgetSize>(size, ignoreCase: true);

    private static JsonObject Expand(string template, string dataJson, string size)
    {
        var expanded = new AdaptiveCardTemplate(template).Expand(new EvaluationContext
        {
            Root = dataJson,
            Host = $$"""{ "widgetSize": "{{size}}" }""",
        });

        Assert.DoesNotContain("${", expanded, StringComparison.Ordinal);
        var node = JsonNode.Parse(expanded)!.AsObject();
        Assert.Equal("AdaptiveCard", node["type"]!.GetValue<string>());
        return node;
    }

    private static IEnumerable<JsonObject> Descendants(JsonNode node)
    {
        switch (node)
        {
            case JsonObject obj:
                yield return obj;
                foreach (var child in obj.Select(p => p.Value).OfType<JsonNode>().SelectMany(Descendants))
                {
                    yield return child;
                }

                break;
            case JsonArray array:
                foreach (var child in array.OfType<JsonNode>().SelectMany(Descendants))
                {
                    yield return child;
                }

                break;
        }
    }

    private static List<string> Texts(JsonObject card) => Descendants(card)
        .Where(n => n["type"]?.GetValue<string>() == "TextBlock")
        .Select(n => n["text"]!.GetValue<string>())
        .ToList();

    private static List<JsonObject> Actions(JsonObject card) => Descendants(card)
        .Where(n => n["type"]?.GetValue<string>() is "Action.OpenUrl" or "Action.Execute")
        .Where(n => n.Parent is JsonArray)
        .ToList();

    private static List<string> ActionTitles(JsonObject card) => Actions(card).Select(a => a["title"]!.GetValue<string>()).ToList();

    private static List<string> SelectActionUrls(JsonObject card) => Descendants(card)
        .Where(n => n["selectAction"] is JsonObject)
        .Select(n => n["selectAction"]!["url"]!.GetValue<string>())
        .ToList();
}
