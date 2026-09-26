using System.Text.Json;
using System.Text.Json.Serialization;

namespace OutlookCalendarWidget.Core;

/// <summary>How far ahead a widget instance looks.</summary>
public enum AgendaRange
{
    Today,
    TwoDays,
    Week,
}

/// <summary>Widget sizes supported by the Windows Widgets Board.</summary>
public enum WidgetSize
{
    Small,
    Medium,
    Large,
}

/// <summary>Per-widget options chosen in the widget's "Customize" pane. Persisted in the widget's custom state.</summary>
public sealed record WidgetPreferences
{
    public static WidgetPreferences Default { get; } = new();

    public AgendaRange Range { get; init; } = AgendaRange.Today;

    public bool ShowDeclined { get; init; }

    public bool ShowAllDay { get; init; } = true;

    public string Serialize() => JsonSerializer.Serialize(this, PreferencesJsonContext.Default.WidgetPreferences);

    public static WidgetPreferences Deserialize(string? customState)
    {
        if (string.IsNullOrWhiteSpace(customState))
        {
            return Default;
        }

        try
        {
            return JsonSerializer.Deserialize(customState, PreferencesJsonContext.Default.WidgetPreferences) ?? Default;
        }
        catch (JsonException)
        {
            return Default;
        }
    }

    /// <summary>Applies the inputs submitted from the customization Adaptive Card (a flat JSON object of string values).</summary>
    public WidgetPreferences WithCustomizationInputs(string? inputsJson)
    {
        if (string.IsNullOrWhiteSpace(inputsJson))
        {
            return this;
        }

        Dictionary<string, string>? inputs;
        try
        {
            inputs = JsonSerializer.Deserialize(inputsJson, PreferencesJsonContext.Default.DictionaryStringString);
        }
        catch (JsonException)
        {
            return this;
        }

        if (inputs is null)
        {
            return this;
        }

        var updated = this;
        if (inputs.TryGetValue("range", out var range) && TryParseRange(range, out var parsedRange))
        {
            updated = updated with { Range = parsedRange };
        }

        if (inputs.TryGetValue("showDeclined", out var showDeclined) && bool.TryParse(showDeclined, out var declined))
        {
            updated = updated with { ShowDeclined = declined };
        }

        if (inputs.TryGetValue("showAllDay", out var showAllDay) && bool.TryParse(showAllDay, out var allDay))
        {
            updated = updated with { ShowAllDay = allDay };
        }

        return updated;
    }

    /// <summary>Builds the data payload bound to the customization template.</summary>
    public string ToCustomizationData() => JsonSerializer.Serialize(
        new CustomizationData(RangeToKey(Range), ShowDeclined ? "true" : "false", ShowAllDay ? "true" : "false"),
        PreferencesJsonContext.Default.CustomizationData);

    internal static string RangeToKey(AgendaRange range) => range switch
    {
        AgendaRange.TwoDays => "twoDays",
        AgendaRange.Week => "week",
        _ => "today",
    };

    private static bool TryParseRange(string value, out AgendaRange range)
    {
        switch (value)
        {
            case "today":
                range = AgendaRange.Today;
                return true;
            case "twoDays":
                range = AgendaRange.TwoDays;
                return true;
            case "week":
                range = AgendaRange.Week;
                return true;
            default:
                range = AgendaRange.Today;
                return false;
        }
    }
}

internal sealed record CustomizationData(string Range, string ShowDeclined, string ShowAllDay);

[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UseStringEnumConverter = true)]
[JsonSerializable(typeof(WidgetPreferences))]
[JsonSerializable(typeof(CustomizationData))]
[JsonSerializable(typeof(Dictionary<string, string>))]
internal sealed partial class PreferencesJsonContext : JsonSerializerContext;
