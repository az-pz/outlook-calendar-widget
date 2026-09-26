using Windows.ApplicationModel;

namespace OutlookCalendarWidget.Services;

/// <summary>
/// Cross-process "please refresh" signal. The UI app sets it after sign-in, sign-out or a settings change so the
/// widget provider process (if running) refreshes immediately instead of waiting for its next timer tick.
/// </summary>
internal static class RefreshSignal
{
    private static readonly string Name = $@"Local\{Package.Current.Id.FamilyName}.Refresh";

    public static void Notify()
    {
        if (EventWaitHandle.TryOpenExisting(Name, out var handle))
        {
            using (handle)
            {
                handle.Set();
            }
        }
    }

    public static IDisposable Listen(Action onSignaled)
    {
        var handle = new EventWaitHandle(false, EventResetMode.AutoReset, Name);
        var registration = ThreadPool.RegisterWaitForSingleObject(handle, (_, _) => onSignaled(), null, Timeout.Infinite, executeOnlyOnce: false);
        return new Subscription(handle, registration);
    }

    private sealed class Subscription(EventWaitHandle handle, RegisteredWaitHandle registration) : IDisposable
    {
        public void Dispose()
        {
            registration.Unregister(null);
            handle.Dispose();
        }
    }
}
