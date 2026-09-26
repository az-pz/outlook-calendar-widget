using System.Runtime.InteropServices;
using Microsoft.Identity.Client;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using OutlookCalendarWidget.Core;
using OutlookCalendarWidget.Services;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;

namespace OutlookCalendarWidget;

/// <summary>Companion window: pick the calendar source, preview today's agenda and sign in to Microsoft Graph if needed.</summary>
public sealed partial class MainWindow : Window
{
    private static readonly TimeSpan OutlookStartupDelay = TimeSpan.FromSeconds(8);
    private readonly CalendarService calendar = CalendarService.Shared;
    private readonly nint hwnd;
    private bool signingIn;
    private bool loadingForm;

    public MainWindow()
    {
        InitializeComponent();
        hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);

        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);
        AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico"));
        var scale = GetDpiForWindow(hwnd) / 96.0;
        AppWindow.Resize(new SizeInt32((int)(620 * scale), (int)(820 * scale)));

        LoadSettingsIntoForm();
        _ = RefreshAsync();
    }

    public void BringToFront()
    {
        if (AppWindow.Presenter is OverlappedPresenter { State: OverlappedPresenterState.Minimized } presenter)
        {
            presenter.Restore();
        }

        AppWindow.Show();
        Activate();
        _ = SetForegroundWindow(hwnd);
    }

    public void ShowSettings()
    {
        if (AppSettings.Load().Source == CalendarSource.Microsoft365)
        {
            SettingsExpander.IsExpanded = true;
            ClientIdBox.Focus(FocusState.Programmatic);
        }

        SourceCard.StartBringIntoView();
    }

    public async Task SignInAsync()
    {
        if (signingIn)
        {
            return;
        }

        var settings = AppSettings.Load();
        if (!settings.IsConfigured)
        {
            ShowSettings();
            ShowStatus(InfoBarSeverity.Warning, "Almost there", "Enter the application (client) ID of your Entra app registration, then sign in.");
            return;
        }

        signingIn = true;
        SignInButton.IsEnabled = false;
        try
        {
            await AuthService.Shared.SignInAsync(settings, hwnd, CancellationToken.None);
            RefreshSignal.Notify();
            await RefreshAsync();
        }
        catch (MsalClientException ex) when (ex.ErrorCode == MsalError.AuthenticationCanceledError)
        {
        }
        catch (Exception ex)
        {
            ShowStatus(InfoBarSeverity.Error, "Sign-in failed", FirstLine(ex.Message));
        }
        finally
        {
            signingIn = false;
            SignInButton.IsEnabled = true;
        }
    }

    private async Task RefreshAsync()
    {
        LoadingBar.Visibility = Visibility.Visible;
        RefreshButton.IsEnabled = false;
        try
        {
            var snapshot = await Task.Run(() => calendar.RefreshAsync());
            Render(snapshot);
        }
        catch (Exception ex)
        {
            ShowStatus(InfoBarSeverity.Error, "Couldn't load your calendar", FirstLine(ex.Message));
        }
        finally
        {
            LoadingBar.Visibility = Visibility.Collapsed;
            RefreshButton.IsEnabled = true;
        }
    }

    private void Render(CalendarSnapshot snapshot)
    {
        StatusInfoBar.IsOpen = false;
        var graph = snapshot.Source == CalendarSource.Microsoft365;
        var outlook = snapshot.Source == CalendarSource.Outlook;
        var ok = snapshot.Status == CalendarStatus.Ok;
        SignInButton.Visibility = graph && !ok ? Visibility.Visible : Visibility.Collapsed;
        SignOutButton.Visibility = graph && ok ? Visibility.Visible : Visibility.Collapsed;
        OpenOutlookButton.Visibility = outlook && snapshot.Status != CalendarStatus.OutlookNotInstalled ? Visibility.Visible : Visibility.Collapsed;
        OpenOutlookButton.Style = (Style)Application.Current.Resources[
            snapshot.Status == CalendarStatus.OutlookNotRunning ? "AccentButtonStyle" : "DefaultButtonStyle"];
        AccountPicture.DisplayName = snapshot.Account ?? string.Empty;

        switch (snapshot.Status)
        {
            case CalendarStatus.OutlookNotRunning:
                AccountTitle.Text = "Outlook isn't running";
                AccountSubtitle.Text = "Open Outlook and the widget picks up your calendar within a minute.";
                break;
            case CalendarStatus.NewOutlookOnly:
                AccountTitle.Text = "Classic Outlook isn't running";
                AccountSubtitle.Text = "The new Outlook doesn't share its calendar with other apps.";
                ShowStatus(InfoBarSeverity.Informational, "Using the new Outlook?", "Open classic Outlook, or connect your Microsoft 365 account instead.");
                StatusInfoBar.ActionButton = SourceButton("Use Microsoft 365", CalendarSource.Microsoft365);
                break;
            case CalendarStatus.OutlookNotInstalled:
                AccountTitle.Text = "Classic Outlook not found";
                AccountSubtitle.Text = string.Empty;
                ShowStatus(InfoBarSeverity.Warning, "Classic Outlook isn't installed", "Connect your Microsoft 365 or Outlook.com account instead.");
                StatusInfoBar.ActionButton = SourceButton("Use Microsoft 365", CalendarSource.Microsoft365);
                break;
            case CalendarStatus.NotConfigured:
                AccountTitle.Text = "Not connected";
                AccountSubtitle.Text = "Connect a Microsoft Entra app registration to get started.";
                ShowStatus(InfoBarSeverity.Informational, "Finish setup", "Enter your app registration's client ID under Microsoft Graph connection, or pick another calendar source.");
                StatusInfoBar.ActionButton = OutlookDesktop.GetAvailability() == OutlookAvailability.NotInstalled
                    ? SourceButton("Use sample events", CalendarSource.Sample)
                    : SourceButton("Use Outlook on this PC", CalendarSource.Outlook);
                SettingsExpander.IsExpanded = true;
                break;
            case CalendarStatus.SignedOut:
                AccountTitle.Text = "Not signed in";
                AccountSubtitle.Text = "Sign in with your work, school or personal Microsoft account.";
                break;
            case CalendarStatus.Error:
                AccountTitle.Text = "Couldn't connect";
                AccountSubtitle.Text = string.Empty;
                ShowStatus(InfoBarSeverity.Error, "Couldn't load your calendar", snapshot.Error ?? string.Empty);
                break;
            case CalendarStatus.Ok when snapshot.Source == CalendarSource.Sample:
                AccountTitle.Text = SampleCalendar.AccountName;
                AccountSubtitle.Text = "Demo mode: the widget shows made-up meetings. Pick another calendar source below.";
                break;
            case CalendarStatus.Ok:
                AccountTitle.Text = snapshot.Account ?? (outlook ? "Outlook" : "Signed in");
                var updated = snapshot.FetchedAt is { } fetched ? $"widget data updated {fetched:t}" : string.Empty;
                AccountSubtitle.Text = outlook ? $"Outlook on this PC · {updated}" : Capitalize(updated);
                if (snapshot.Error is not null)
                {
                    ShowStatus(InfoBarSeverity.Warning, "Showing cached events", snapshot.Error);
                }

                break;
        }

        RenderAgenda(snapshot);
    }

    private void RenderAgenda(CalendarSnapshot snapshot)
    {
        if (snapshot.Status != CalendarStatus.Ok)
        {
            AgendaList.ItemsSource = null;
            AgendaSummary.Text = string.Empty;
            AgendaEmpty.Text = snapshot.Status switch
            {
                CalendarStatus.Error => "Try again in a moment.",
                CalendarStatus.Loading => "Loading your calendar…",
                CalendarStatus.OutlookNotRunning or CalendarStatus.NewOutlookOnly => "Open Outlook to see your meetings.",
                CalendarStatus.OutlookNotInstalled => "Choose a calendar source below.",
                _ => "Sign in to see your meetings.",
            };
            AgendaEmpty.Visibility = Visibility.Visible;
            return;
        }

        var card = calendar.BuildCard(WidgetPreferences.Default, WidgetSize.Large, maxEvents: 12);
        var items = card.Days.SelectMany(d => d.Events).Select(AgendaItem.From).ToList();
        AgendaList.ItemsSource = items;
        AgendaSummary.Text = $"{card.DateLabel} · {card.Summary}";
        AgendaEmpty.Text = $"{card.EmptyTitle}. {card.EmptyText}";
        AgendaEmpty.Visibility = items.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void LoadSettingsIntoForm()
    {
        var settings = AppSettings.Load();
        loadingForm = true;
        SourceButtons.SelectedIndex = (int)settings.Source;
        loadingForm = false;
        SettingsExpander.Visibility = settings.Source == CalendarSource.Microsoft365 ? Visibility.Visible : Visibility.Collapsed;
        ClientIdBox.Text = settings.ClientId;
        TenantBox.Text = settings.TenantId;
        OutlookUrlBox.Text = settings.OutlookUrl;
        UpdateRedirectUri();
    }

    private Button SourceButton(string title, CalendarSource source)
    {
        var button = new Button { Content = title };
        button.Click += (_, _) => SourceButtons.SelectedIndex = (int)source;
        return button;
    }

    private void UpdateRedirectUri()
    {
        var clientId = ClientIdBox.Text.Trim();
        RedirectUriBox.Text = Guid.TryParse(clientId, out _) ? $"ms-appx-web://microsoft.aad.brokerplugin/{clientId}" : string.Empty;
    }

    private void ShowStatus(InfoBarSeverity severity, string title, string message)
    {
        StatusInfoBar.ActionButton = null;
        StatusInfoBar.Severity = severity;
        StatusInfoBar.Title = title;
        StatusInfoBar.Message = message;
        StatusInfoBar.IsOpen = true;
    }

    private async void OnSignInClick(object sender, RoutedEventArgs e) => await SignInAsync();

    private async void OnSignOutClick(object sender, RoutedEventArgs e)
    {
        try
        {
            await AuthService.Shared.SignOutAsync(AppSettings.Load());
            RefreshSignal.Notify();
            await RefreshAsync();
        }
        catch (Exception ex)
        {
            ShowStatus(InfoBarSeverity.Error, "Sign-out failed", FirstLine(ex.Message));
        }
    }

    private async void OnRefreshClick(object sender, RoutedEventArgs e)
    {
        RefreshSignal.Notify();
        await RefreshAsync();
    }

    private void OnClientIdChanged(object sender, TextChangedEventArgs e) => UpdateRedirectUri();

    private async void OnSourceChanged(object sender, SelectionChangedEventArgs e)
    {
        if (loadingForm || SourceButtons.SelectedIndex < 0)
        {
            return;
        }

        var source = (CalendarSource)SourceButtons.SelectedIndex;
        (AppSettings.Load() with { Source = source }).Save();
        SettingsExpander.Visibility = source == CalendarSource.Microsoft365 ? Visibility.Visible : Visibility.Collapsed;
        RefreshSignal.Notify();
        await RefreshAsync();
    }

    private async void OnOpenOutlookClick(object sender, RoutedEventArgs e)
    {
        var wasRunning = OutlookDesktop.GetAvailability() == OutlookAvailability.Running;
        try
        {
            await Task.Run(() => OutlookDesktop.Open(new OutlookOpenRequest(null, null)));
        }
        catch (Exception ex)
        {
            ShowStatus(InfoBarSeverity.Error, "Couldn't open Outlook", FirstLine(ex.Message));
            return;
        }

        if (!wasRunning)
        {
            // Give Outlook a moment to load its profile, then pick up the calendar (the widget polls on its own too).
            await Task.Delay(OutlookStartupDelay);
            RefreshSignal.Notify();
            await RefreshAsync();
        }
    }

    private void OnCopyRedirectUriClick(object sender, RoutedEventArgs e)
    {
        if (RedirectUriBox.Text.Length == 0)
        {
            return;
        }

        var package = new DataPackage();
        package.SetText(RedirectUriBox.Text);
        Clipboard.SetContent(package);
    }

    private async void OnSaveSettingsClick(object sender, RoutedEventArgs e)
    {
        var clientId = ClientIdBox.Text.Trim();
        if (clientId.Length > 0 && !Guid.TryParse(clientId, out _))
        {
            ShowStatus(InfoBarSeverity.Error, "Invalid client ID", "The application (client) ID must be a GUID.");
            return;
        }

        var outlookUrl = OutlookUrlBox.Text.Trim();
        if (!Uri.TryCreate(outlookUrl, UriKind.Absolute, out _))
        {
            outlookUrl = AppLinks.DefaultOutlookCalendarUrl;
        }

        var tenant = TenantBox.Text.Trim();
        (AppSettings.Load() with
        {
            ClientId = clientId,
            TenantId = tenant.Length == 0 ? "common" : tenant,
            OutlookUrl = outlookUrl,
        }).Save();

        LoadSettingsIntoForm();
        RefreshSignal.Notify();
        await RefreshAsync();
        if (StatusInfoBar is not { IsOpen: true, Severity: InfoBarSeverity.Error })
        {
            ShowStatus(InfoBarSeverity.Success, "Saved", "Your widget will use the new settings.");
        }
    }

    private static string FirstLine(string message) => message.Split('\n', 2)[0].Trim();

    private static string Capitalize(string text) => text.Length == 0 ? text : char.ToUpperInvariant(text[0]) + text[1..];

    [LibraryImport("user32.dll")]
    private static partial uint GetDpiForWindow(nint hwnd);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SetForegroundWindow(nint hwnd);
}

/// <summary>Row model for the agenda preview.</summary>
public sealed class AgendaItem
{
    public required string Subject { get; init; }

    public required string Time { get; init; }

    public required string End { get; init; }

    public required string Detail { get; init; }

    public required Brush Accent { get; init; }

    internal static AgendaItem From(CardEvent e) => new()
    {
        Subject = e.Subject,
        Time = e.TimeText,
        End = e.EndText,
        Detail = string.Join(" · ", new[] { e.RelativeText, e.Detail }.Where(s => s.Length > 0)),
        Accent = ToBrush(StatusBar.Color(e.StatusStyle)),
    };

    private static SolidColorBrush ToBrush(uint argb) =>
        new(ColorHelper.FromArgb((byte)(argb >> 24), (byte)(argb >> 16), (byte)(argb >> 8), (byte)argb));
}
