using Microsoft.UI.Xaml;
using Microsoft.Windows.AppLifecycle;
using OutlookCalendarWidget.Core;
using Windows.ApplicationModel.Activation;

namespace OutlookCalendarWidget;

public partial class App : Application
{
    private readonly AppActivationArguments initialActivation;
    private MainWindow? window;

    public App(AppActivationArguments activation)
    {
        initialActivation = activation;
        InitializeComponent();
        AppInstance.GetCurrent().Activated += OnRedirectedActivation;
    }

    protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
    {
        window = new MainWindow();
        window.Activate();
        HandleActivation(initialActivation);
    }

    private void OnRedirectedActivation(object? sender, AppActivationArguments args)
    {
        window?.DispatcherQueue.TryEnqueue(() =>
        {
            window.BringToFront();
            HandleActivation(args);
        });
    }

    private void HandleActivation(AppActivationArguments args)
    {
        var target = args.Kind switch
        {
            ExtendedActivationKind.Protocol when args.Data is IProtocolActivatedEventArgs protocol => protocol.Uri.OriginalString,
            ExtendedActivationKind.Launch when args.Data is ILaunchActivatedEventArgs launch => launch.Arguments,
            _ => string.Empty,
        };

        if (target.Contains("signin", StringComparison.OrdinalIgnoreCase))
        {
            _ = window?.SignInAsync();
        }
        else if (target.StartsWith(AppLinks.Settings, StringComparison.OrdinalIgnoreCase))
        {
            window?.ShowSettings();
        }
    }
}
