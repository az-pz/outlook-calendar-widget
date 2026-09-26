using System.Collections.Concurrent;
using Microsoft.Windows.Widgets.Providers;
using OutlookCalendarWidget.Core;
using OutlookCalendarWidget.Services;
using CoreWidgetSize = OutlookCalendarWidget.Core.WidgetSize;
using HostWidgetSize = Microsoft.Windows.Widgets.WidgetSize;

namespace OutlookCalendarWidget.Widgets;

/// <summary>
/// The Windows Widgets Board calls into this object (over COM) to create, update, resize, customize and remove
/// Outlook Calendar widgets. All pinned widget instances share one calendar cache.
/// </summary>
[System.Diagnostics.CodeAnalysis.SuppressMessage("Design", "CA1001:Types that own disposable fields should be disposable", Justification = "Process-lifetime singleton.")]
internal sealed partial class CalendarWidgetProvider : IWidgetProvider, IWidgetProvider2
{
    public const string AgendaDefinitionId = "OutlookCalendar_Agenda";

    private static readonly TimeSpan TickInterval = TimeSpan.FromMinutes(1);
    private static readonly TimeSpan ActiveRefreshInterval = TimeSpan.FromMinutes(5);
    private static readonly TimeSpan BackgroundRefreshInterval = TimeSpan.FromMinutes(15);

    // Reading the local Outlook calendar costs no network round trip, so new or moved meetings can show up sooner.
    private static readonly TimeSpan LocalActiveRefreshInterval = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan LocalBackgroundRefreshInterval = TimeSpan.FromMinutes(5);

    private static readonly Lazy<CalendarWidgetProvider> LazyInstance = new(() => new CalendarWidgetProvider());

    private readonly ConcurrentDictionary<string, WidgetInstance> widgets = new();
    private readonly CalendarService calendar = CalendarService.Shared;
    private readonly Timer timer;
    private readonly IDisposable refreshSubscription;
    private int refreshInFlight;
    private long lastRefreshTicks;

    private CalendarWidgetProvider()
    {
        // Recover widgets that were pinned before this process (re)started. Returns null (not empty) when there are none.
        foreach (var info in WidgetManager.GetDefault().GetWidgetInfos() ?? [])
        {
            var context = info.WidgetContext;
            widgets[context.Id] = new WidgetInstance(context.Id)
            {
                Size = Map(context.Size),
                IsActive = context.IsActive,
                Preferences = WidgetPreferences.Deserialize(info.CustomState),
            };
        }

        calendar.SnapshotChanged += (_, _) => RenderAll();
        refreshSubscription = RefreshSignal.Listen(QueueRefresh);
        timer = new Timer(_ => OnTick(), null, TickInterval, TickInterval);

        Log.Info($"Widget provider started with {widgets.Count} existing widget(s){(widgets.IsEmpty ? string.Empty : ": " + string.Join(", ", widgets.Values.Select(w => $"{w.Id} ({w.Size})")))}.");
        if (!widgets.IsEmpty)
        {
            QueueRefresh();
        }
    }

    public static CalendarWidgetProvider Instance => LazyInstance.Value;

    /// <summary>Signaled when the last widget is removed so the COM server process can exit.</summary>
    public static ManualResetEvent NoWidgetsRemaining { get; } = new(false);

    public void CreateWidget(WidgetContext widgetContext)
    {
        var widget = new WidgetInstance(widgetContext.Id) { Size = Map(widgetContext.Size), IsActive = widgetContext.IsActive };
        widgets[widget.Id] = widget;
        NoWidgetsRemaining.Reset();
        Log.Info($"Widget {widget.Id} created ({widget.Size}).");

        Render(widget);
        if (calendar.Snapshot.Status != CalendarStatus.Ok || calendar.Age > ActiveRefreshInterval)
        {
            QueueRefresh();
        }
    }

    public void DeleteWidget(string widgetId, string customState)
    {
        widgets.TryRemove(widgetId, out _);
        Log.Info($"Widget {widgetId} deleted; {widgets.Count} remaining.");
        if (widgets.IsEmpty)
        {
            NoWidgetsRemaining.Set();
        }
    }

    public void OnActionInvoked(WidgetActionInvokedArgs actionInvokedArgs)
    {
        var widgetId = actionInvokedArgs.WidgetContext.Id;
        var verb = actionInvokedArgs.Verb;
        var data = actionInvokedArgs.Data;
        if (!widgets.TryGetValue(widgetId, out var widget))
        {
            return;
        }

        switch (verb)
        {
            case WidgetVerbs.Refresh:
                QueueRefresh();
                break;
            case WidgetVerbs.SaveCustomization:
                widget.Preferences = widget.Preferences.WithCustomizationInputs(data);
                widget.InCustomization = false;
                Render(widget);
                break;
            case WidgetVerbs.ExitCustomization:
                widget.InCustomization = false;
                Render(widget);
                break;
        }
    }

    public void OnWidgetContextChanged(WidgetContextChangedArgs contextChangedArgs)
    {
        var context = contextChangedArgs.WidgetContext;
        if (widgets.TryGetValue(context.Id, out var widget))
        {
            widget.Size = Map(context.Size);
            Log.Info($"Widget {widget.Id} resized ({widget.Size}).");
            Render(widget);
        }
    }

    public void Activate(WidgetContext widgetContext)
    {
        if (widgets.TryGetValue(widgetContext.Id, out var widget))
        {
            widget.IsActive = true;
            Render(widget);
            if (calendar.Age > TimeSpan.FromMinutes(2))
            {
                QueueRefresh();
            }
        }
    }

    public void Deactivate(string widgetId)
    {
        if (widgets.TryGetValue(widgetId, out var widget))
        {
            widget.IsActive = false;
        }
    }

    public void OnCustomizationRequested(WidgetCustomizationRequestedArgs customizationRequestedArgs)
    {
        if (widgets.TryGetValue(customizationRequestedArgs.WidgetContext.Id, out var widget))
        {
            widget.InCustomization = true;
            Render(widget);
        }
    }

    private void OnTick()
    {
        var anyActive = widgets.Values.Any(w => w.IsActive);
        var snapshot = calendar.Snapshot;
        var threshold = snapshot switch
        {
            { RetrySoon: true } => TickInterval,
            { Source: CalendarSource.Outlook } => anyActive ? LocalActiveRefreshInterval : LocalBackgroundRefreshInterval,
            _ => anyActive ? ActiveRefreshInterval : BackgroundRefreshInterval,
        };
        var sinceLastRefresh = TimeSpan.FromTicks(DateTime.UtcNow.Ticks - Interlocked.Read(ref lastRefreshTicks));
        if (sinceLastRefresh >= threshold)
        {
            QueueRefresh();
        }
        else
        {
            // Nothing to fetch, but relative times ("in 5 min") and the "now" highlight still need to move.
            RenderAll();
        }
    }

    private void QueueRefresh()
    {
        if (Interlocked.Exchange(ref refreshInFlight, 1) == 1)
        {
            return;
        }

        Interlocked.Exchange(ref lastRefreshTicks, DateTime.UtcNow.Ticks);

        _ = Task.Run(async () =>
        {
            try
            {
                await calendar.RefreshAsync();
            }
            catch (Exception ex)
            {
                Log.Error("Calendar refresh failed", ex);
            }
            finally
            {
                Volatile.Write(ref refreshInFlight, 0);
            }
        });
    }

    private void RenderAll()
    {
        foreach (var widget in widgets.Values)
        {
            Render(widget);
        }
    }

    private void Render(WidgetInstance widget)
    {
        try
        {
            var options = new WidgetUpdateRequestOptions(widget.Id);
            if (widget.InCustomization)
            {
                options.Template = WidgetTemplates.Customize;
                options.Data = widget.Preferences.ToCustomizationData();
            }
            else
            {
                options.Template = WidgetTemplates.Calendar;
                options.Data = calendar.BuildCard(widget.Preferences, widget.Size).ToJson();
            }

            options.CustomState = widget.Preferences.Serialize();
            WidgetManager.GetDefault().UpdateWidget(options);
        }
        catch (Exception ex)
        {
            // The widget may have been removed between the lookup and the update.
            Log.Error($"Failed to update widget {widget.Id}", ex);
        }
    }

    private static CoreWidgetSize Map(HostWidgetSize size) => size switch
    {
        HostWidgetSize.Small => CoreWidgetSize.Small,
        HostWidgetSize.Large => CoreWidgetSize.Large,
        _ => CoreWidgetSize.Medium,
    };

    private sealed class WidgetInstance(string id)
    {
        public string Id { get; } = id;

        public CoreWidgetSize Size { get; set; } = CoreWidgetSize.Medium;

        public bool IsActive { get; set; }

        public bool InCustomization { get; set; }

        public WidgetPreferences Preferences { get; set; } = WidgetPreferences.Default;
    }
}
