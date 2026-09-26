using OutlookCalendarWidget.Core;

namespace OutlookCalendarWidget.Core.Tests;

public class WidgetPreferencesTests
{
    [Fact]
    public void Round_trips_through_custom_state()
    {
        var preferences = new WidgetPreferences { Range = AgendaRange.Week, ShowDeclined = true, ShowAllDay = false };

        var restored = WidgetPreferences.Deserialize(preferences.Serialize());

        Assert.Equal(preferences, restored);
        Assert.Contains("\"range\":\"Week\"", preferences.Serialize(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("not json")]
    [InlineData("42")]
    public void Invalid_custom_state_falls_back_to_defaults(string? customState)
    {
        Assert.Equal(WidgetPreferences.Default, WidgetPreferences.Deserialize(customState));
    }

    [Fact]
    public void Applies_customization_inputs()
    {
        var updated = WidgetPreferences.Default.WithCustomizationInputs(
            """{ "range": "twoDays", "showDeclined": "true", "showAllDay": "false" }""");

        Assert.Equal(new WidgetPreferences { Range = AgendaRange.TwoDays, ShowDeclined = true, ShowAllDay = false }, updated);
    }

    [Fact]
    public void Ignores_unknown_or_malformed_inputs()
    {
        var original = new WidgetPreferences { Range = AgendaRange.Week };

        Assert.Equal(original, original.WithCustomizationInputs("""{ "range": "month", "showDeclined": "maybe" }"""));
        Assert.Equal(original, original.WithCustomizationInputs("garbage"));
        Assert.Equal(original, original.WithCustomizationInputs(null));
    }

    [Fact]
    public void Customization_data_uses_template_keys()
    {
        var data = new WidgetPreferences { Range = AgendaRange.TwoDays, ShowDeclined = true }.ToCustomizationData();

        Assert.Equal("""{"range":"twoDays","showDeclined":"true","showAllDay":"true"}""", data);
    }
}
