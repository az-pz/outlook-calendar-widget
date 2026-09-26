using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.Windows.AppLifecycle;
using OutlookCalendarWidget.Core;
using OutlookCalendarWidget.Services;
using OutlookCalendarWidget.Widgets;
using Windows.ApplicationModel.Activation;

namespace OutlookCalendarWidget;

/// <summary>
/// One executable, two roles:
/// <list type="bullet">
/// <item>Started by COM with <c>-RegisterProcessAsComServer</c> (see Package.appxmanifest): runs headless as the widget provider.</item>
/// <item>Started from a widget tap on a meeting (<c>outlook-calendar-widget:open</c>): shows it in Outlook and exits.</item>
/// <item>Started from Start, a widget button (protocol link) or the command line: shows the sign-in / settings window.</item>
/// </list>
/// </summary>
public static class Program
{
    private const string ComServerSwitch = "-RegisterProcessAsComServer";

    [STAThread]
    public static int Main(string[] args)
    {
        WinRT.ComWrappersSupport.InitializeComWrappers();

        if (args.Any(a => string.Equals(a, ComServerSwitch, StringComparison.OrdinalIgnoreCase)))
        {
            WidgetProviderServer.Run();
            return 0;
        }

        // Single-instance UI: forward activations (for example "Sign in" from a widget) to the running window.
        var activation = AppInstance.GetCurrent().GetActivatedEventArgs();
        if (TryOpenInOutlook(activation))
        {
            return 0;
        }

        var mainInstance = AppInstance.FindOrRegisterForKey("main");
        if (!mainInstance.IsCurrent)
        {
            Task.Run(() => mainInstance.RedirectActivationToAsync(activation).AsTask()).Wait();
            return 0;
        }

        Application.Start(callbackParams =>
        {
            SynchronizationContext.SetSynchronizationContext(new DispatcherQueueSynchronizationContext(DispatcherQueue.GetForCurrentThread()));
            _ = new App(activation);
        });

        return 0;
    }

    /// <summary>
    /// Widget taps on a meeting (or "Open Outlook") arrive as <c>outlook-calendar-widget:open…</c> links. They are
    /// handled here without any UI: show the item in classic Outlook and exit.
    /// </summary>
    private static bool TryOpenInOutlook(AppActivationArguments activation)
    {
        if (activation.Kind != ExtendedActivationKind.Protocol ||
            activation.Data is not IProtocolActivatedEventArgs protocol ||
            !AppLinks.TryParseOpenOutlook(protocol.Uri.OriginalString, out var request))
        {
            return false;
        }

        try
        {
            OutlookDesktop.Open(request);
        }
        catch (Exception ex)
        {
            Log.Error("Couldn't open Outlook", ex);
        }

        return true;
    }
}
