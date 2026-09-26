namespace OutlookCalendarWidget.Core;

/// <summary>Adaptive Card templates sent to the Widgets Board. Embedded in this assembly from the Templates folder.</summary>
public static class WidgetTemplates
{
    private static readonly Lazy<string> CalendarTemplate = new(() => Load("CalendarWidget.json"));
    private static readonly Lazy<string> CustomizeTemplate = new(() => Load("CustomizeWidget.json"));

    public static string Calendar => CalendarTemplate.Value;

    public static string Customize => CustomizeTemplate.Value;

    private static string Load(string name)
    {
        var resourceName = $"OutlookCalendarWidget.Core.Templates.{name}";
        using var stream = typeof(WidgetTemplates).Assembly.GetManifestResourceStream(resourceName)
                           ?? throw new InvalidOperationException($"Missing embedded template '{resourceName}'.");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
