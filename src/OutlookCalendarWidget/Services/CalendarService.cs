using System.Globalization;
using Microsoft.Identity.Client;
using OutlookCalendarWidget.Core;

namespace OutlookCalendarWidget.Services;

internal enum CalendarStatus
{
    Loading,
    NotConfigured,
    SignedOut,
    OutlookNotRunning,
    NewOutlookOnly,
    OutlookNotInstalled,
    Ok,
    Error,
}

internal sealed record CalendarSnapshot(
    CalendarStatus Status,
    IReadOnlyList<CalendarEvent> Events,
    DateTimeOffset? FetchedAt,
    string? Error = null,
    string? Account = null)
{
    public static CalendarSnapshot Initial { get; } = new(CalendarStatus.Loading, [], null);

    public CalendarSource Source { get; init; }

    /// <summary>Local Outlook reads are cheap, so poll every minute until Outlook is back (opened, or no longer busy).</summary>
    public bool RetrySoon => Source == CalendarSource.Outlook && (Status != CalendarStatus.Ok || Error is not null);
}

/// <summary>Fetches and caches the user's upcoming events and turns them into widget card data.</summary>
[System.Diagnostics.CodeAnalysis.SuppressMessage("Design", "CA1001:Types that own disposable fields should be disposable", Justification = "Process-lifetime singleton.")]
internal sealed class CalendarService
{
    private readonly SemaphoreSlim refreshGate = new(1, 1);
    private readonly HttpClient httpClient = new() { Timeout = TimeSpan.FromSeconds(30) };

    public static CalendarService Shared { get; } = new();

    public CalendarSnapshot Snapshot { get; private set; } = CalendarSnapshot.Initial;

    public event EventHandler? SnapshotChanged;

    public TimeSpan Age => Snapshot.FetchedAt is { } fetched ? DateTimeOffset.Now - fetched : TimeSpan.MaxValue;

    public async Task<CalendarSnapshot> RefreshAsync(CancellationToken cancellationToken = default)
    {
        await refreshGate.WaitAsync(cancellationToken);
        try
        {
            var previous = Snapshot;
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            Snapshot = await FetchAsync(cancellationToken);
            LogIfChanged(previous, Snapshot, stopwatch.Elapsed);
        }
        finally
        {
            refreshGate.Release();
        }

        SnapshotChanged?.Invoke(this, EventArgs.Empty);
        return Snapshot;
    }

    private static void LogIfChanged(CalendarSnapshot previous, CalendarSnapshot current, TimeSpan elapsed)
    {
        if (previous.Status == current.Status && previous.Source == current.Source &&
            previous.Error == current.Error && previous.Events.Count == current.Events.Count)
        {
            return;
        }

        Log.Info($"Calendar ({current.Source}): {current.Status}, {current.Events.Count} event(s) in {elapsed.TotalMilliseconds:0} ms" +
                 (current.Error is null ? "." : $"; {current.Error}"));
    }

    public CardData BuildCard(WidgetPreferences preferences, WidgetSize size, int? maxEvents = null)
    {
        var snapshot = Snapshot;
        return snapshot.Status switch
        {
            CalendarStatus.OutlookNotRunning => CardData.ForStatus(
                CardStates.Unavailable,
                "Open Outlook to see your meetings",
                "The widget reads your calendar from Outlook on this PC.",
                "Open Outlook",
                actionUrl: AppLinks.OpenOutlook),
            CalendarStatus.NewOutlookOnly => CardData.ForStatus(
                CardStates.Unavailable,
                "Classic Outlook isn't running",
                "The new Outlook doesn't share its calendar with other apps. Open classic Outlook, or connect Microsoft 365 in the app.",
                "Open app",
                actionUrl: AppLinks.Settings),
            CalendarStatus.OutlookNotInstalled => CardData.ForStatus(
                CardStates.Unavailable,
                "Classic Outlook isn't installed",
                "Open the app to connect your Microsoft 365 or Outlook.com account instead.",
                "Open app",
                actionUrl: AppLinks.Settings),
            CalendarStatus.NotConfigured => CardData.ForStatus(
                CardStates.NotConfigured,
                "Connect your calendar",
                "Open the app to finish setting up your Microsoft account.",
                "Open app",
                actionUrl: AppLinks.Settings),
            CalendarStatus.SignedOut => CardData.ForStatus(
                CardStates.SignedOut,
                "Sign in to see your meetings",
                "Connect your Microsoft 365 or Outlook.com account.",
                "Sign in",
                actionUrl: AppLinks.SignIn),
            CalendarStatus.Error => CardData.ForStatus(
                CardStates.Error,
                "Couldn't load your calendar",
                snapshot.Error ?? "Something went wrong.",
                "Retry",
                actionVerb: WidgetVerbs.Refresh),
            CalendarStatus.Ok => AgendaBuilder.Build(
                snapshot.Events,
                DateTimeOffset.Now,
                preferences,
                size,
                TimeZoneInfo.Local,
                CultureInfo.CurrentCulture,
                snapshot.Source == CalendarSource.Outlook ? AppLinks.OpenOutlook : AppSettings.Load().OutlookUrl,
                snapshot.FetchedAt,
                maxEvents),
            _ => CardData.ForStatus(CardStates.Loading, "Loading your calendar…", string.Empty),
        };
    }

    private async Task<CalendarSnapshot> FetchAsync(CancellationToken cancellationToken)
    {
        var settings = AppSettings.Load();
        var snapshot = settings.Source switch
        {
            CalendarSource.Sample => FetchSample(),
            CalendarSource.Microsoft365 => await FetchFromGraphAsync(settings, cancellationToken),
            _ => await FetchFromOutlookAsync(cancellationToken),
        };
        return snapshot with { Source = settings.Source };
    }

    private static CalendarSnapshot FetchSample()
    {
        var now = DateTimeOffset.Now;
        return new CalendarSnapshot(CalendarStatus.Ok, SampleCalendar.Create(now, TimeZoneInfo.Local), now, Account: SampleCalendar.AccountName);
    }

    private async Task<CalendarSnapshot> FetchFromOutlookAsync(CancellationToken cancellationToken)
    {
        switch (OutlookDesktop.GetAvailability())
        {
            case OutlookAvailability.NotInstalled:
                return new CalendarSnapshot(CalendarStatus.OutlookNotInstalled, [], null);
            case OutlookAvailability.NewOutlookOnly:
                return KeepLastGood(CalendarSource.Outlook, "Classic Outlook isn't running.", new CalendarSnapshot(CalendarStatus.NewOutlookOnly, [], null));
            case OutlookAvailability.NotRunning:
                return KeepLastGood(CalendarSource.Outlook, "Outlook isn't running.", new CalendarSnapshot(CalendarStatus.OutlookNotRunning, [], null));
        }

        try
        {
            TimeZoneInfo.ClearCachedData();
            var timeZone = TimeZoneInfo.Local;
            var now = DateTimeOffset.Now;
            var (start, end) = AgendaBuilder.FetchWindow(now, timeZone);

            var calendar = await OutlookDesktop.ReadCalendarAsync(start, end, timeZone, cancellationToken);
            return new CalendarSnapshot(CalendarStatus.Ok, calendar.Events, now, Account: calendar.Account);
        }
        catch (OutlookUnavailableException ex)
        {
            Log.Error("Reading the Outlook calendar failed", ex);
            return KeepLastGood(CalendarSource.Outlook, ex.Message, new CalendarSnapshot(CalendarStatus.Error, [], null, ex.Message));
        }
    }

    private async Task<CalendarSnapshot> FetchFromGraphAsync(AppSettings settings, CancellationToken cancellationToken)
    {
        if (!settings.IsConfigured)
        {
            return new CalendarSnapshot(CalendarStatus.NotConfigured, [], null);
        }

        try
        {
            var auth = await AuthService.Shared.TryAcquireTokenSilentAsync(settings, cancellationToken);
            if (auth is null)
            {
                return new CalendarSnapshot(CalendarStatus.SignedOut, [], null);
            }

            TimeZoneInfo.ClearCachedData();
            var timeZone = TimeZoneInfo.Local;
            var now = DateTimeOffset.Now;
            var (start, end) = AgendaBuilder.FetchWindow(now, timeZone);

            var client = new GraphCalendarClient(httpClient, _ => Task.FromResult(auth.AccessToken));
            var events = await client.GetEventsAsync(start, end, timeZone, cancellationToken);
            return new CalendarSnapshot(CalendarStatus.Ok, events, now, Account: auth.Account?.Username);
        }
        catch (GraphAuthenticationException)
        {
            return new CalendarSnapshot(CalendarStatus.SignedOut, [], null);
        }
        catch (Exception ex) when (ex is not OperationCanceledException || !cancellationToken.IsCancellationRequested)
        {
            var message = ex switch
            {
                HttpRequestException => "Can't reach Microsoft Graph. Check your internet connection.",
                TaskCanceledException => "Microsoft Graph took too long to respond.",
                MsalServiceException msal => FirstLine(msal.Message),
                _ => FirstLine(ex.Message),
            };

            return KeepLastGood(CalendarSource.Microsoft365, message, new CalendarSnapshot(CalendarStatus.Error, [], null, message));
        }
    }

    /// <summary>
    /// Keeps showing the last good agenda from the same source (its "Updated" time reveals it is stale) rather than
    /// replacing it with a status card when a refresh fails.
    /// </summary>
    private CalendarSnapshot KeepLastGood(CalendarSource source, string message, CalendarSnapshot fallback)
    {
        var previous = Snapshot;
        return previous.Status == CalendarStatus.Ok && previous.Source == source
            ? previous with { Error = message }
            : fallback;
    }

    private static string FirstLine(string message)
    {
        var line = message.Split('\n', 2)[0].Trim();
        return line.Length > 200 ? line[..200] + "…" : line;
    }
}

internal static class WidgetVerbs
{
    public const string Refresh = "refresh";
    public const string SaveCustomization = "saveCustomization";
    public const string ExitCustomization = "exitCustomization";
}
