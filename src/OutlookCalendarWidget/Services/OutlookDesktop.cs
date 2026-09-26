using System.Diagnostics;
using System.Runtime.InteropServices;
using OutlookCalendarWidget.Core;

namespace OutlookCalendarWidget.Services;

internal enum OutlookAvailability
{
    Running,
    NotRunning,

    /// <summary>Only the new Outlook (olk.exe) is open; it has no object model other apps can read.</summary>
    NewOutlookOnly,
    NotInstalled,
}

/// <summary>Classic Outlook is missing, closed, busy or not answering.</summary>
internal sealed class OutlookUnavailableException(string message, Exception? innerException = null) : Exception(message, innerException);

internal sealed record OutlookCalendar(IReadOnlyList<CalendarEvent> Events, string? Account);

/// <summary>
/// Reads the default calendar of the classic Outlook for Windows running on this PC through its COM object model
/// (late-bound, so no Office interop assemblies are needed). Only properties outside the Outlook object model
/// guard are read, so this never triggers Outlook's "a program is trying to access…" prompt.
/// </summary>
internal static partial class OutlookDesktop
{
    private const string ProgId = "Outlook.Application";
    private const int MaxItems = 500;
    private static readonly TimeSpan ReadTimeout = TimeSpan.FromSeconds(45);

    public static OutlookAvailability GetAvailability()
    {
        if (Type.GetTypeFromProgID(ProgId) is null)
        {
            return OutlookAvailability.NotInstalled;
        }

        if (IsRunning("OUTLOOK"))
        {
            return OutlookAvailability.Running;
        }

        return IsRunning("olk") ? OutlookAvailability.NewOutlookOnly : OutlookAvailability.NotRunning;
    }

    public static async Task<OutlookCalendar> ReadCalendarAsync(
        DateTimeOffset start,
        DateTimeOffset end,
        TimeZoneInfo timeZone,
        CancellationToken cancellationToken)
    {
        // Outlook stops answering COM calls while it shows some modal dialogs; never let that wedge the refresh loop.
        try
        {
            return await Task.Run(() => ReadCalendar(start, end, timeZone), cancellationToken).WaitAsync(ReadTimeout, cancellationToken);
        }
        catch (TimeoutException ex)
        {
            throw new OutlookUnavailableException("Outlook isn't responding. Close any open Outlook dialogs.", ex);
        }
    }

    /// <summary>Shows an item (or the calendar) in Outlook, starting Outlook normally if it isn't running.</summary>
    public static void Open(OutlookOpenRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (!IsRunning("OUTLOOK"))
        {
            // Started through COM, Outlook would run without a main window, so launch it the normal way instead.
            Process.Start(new ProcessStartInfo("outlook.exe", "/select outlook:calendar") { UseShellExecute = true })?.Dispose();
            return;
        }

        AllowOutlookToComeToFront();
        using var com = new ComScope();
        try
        {
            dynamic app = com.Track(CreateApplication());
            dynamic session = com.Track(app.GetNamespace("MAPI"));
            if (request.EntryId is not null && TryDisplayItem(com, session, request))
            {
                return;
            }

            dynamic calendar = com.Track(session.GetDefaultFolder(OutlookDesktopCalendar.CalendarFolderId));
            object? explorer = app.ActiveExplorer();
            if (explorer is null)
            {
                calendar.Display();
                return;
            }

            dynamic window = com.Track(explorer);
            window.CurrentFolder = calendar;
            window.Activate();
        }
        catch (COMException ex)
        {
            throw Translate(ex);
        }
    }

    private static OutlookCalendar ReadCalendar(DateTimeOffset start, DateTimeOffset end, TimeZoneInfo timeZone)
    {
        using var com = new ComScope();
        try
        {
            dynamic app = com.Track(CreateApplication());
            dynamic session = com.Track(app.GetNamespace("MAPI"));
            dynamic folder = com.Track(session.GetDefaultFolder(OutlookDesktopCalendar.CalendarFolderId));
            string? account = null;
            try
            {
                dynamic store = com.Track(folder.Store);
                account = store.DisplayName as string;
            }
            catch (COMException)
            {
            }

            // Order matters: sort by start, then include recurrences, then restrict (per the Outlook docs).
            dynamic items = com.Track(folder.Items);
            items.Sort("[Start]");
            items.IncludeRecurrences = true;
            dynamic matches = com.Track(items.Restrict(OutlookDesktopCalendar.RestrictFilter(start, end, timeZone)));

            // Count is meaningless once recurrences are included, so walk the collection instead.
            var appointments = new List<OutlookAppointment>();
            object? current = matches.GetFirst();
            while (current is not null && appointments.Count < MaxItems)
            {
                try
                {
                    if (ReadAppointment(current) is { } appointment)
                    {
                        appointments.Add(appointment);
                    }
                }
                finally
                {
                    Release(current);
                }

                current = matches.GetNext();
            }

            return new OutlookCalendar(OutlookDesktopCalendar.ToEvents(appointments, start, end, timeZone), account);
        }
        catch (COMException ex)
        {
            throw Translate(ex);
        }
    }

    private static OutlookAppointment? ReadAppointment(object comItem)
    {
        dynamic item = comItem;
        if ((int)item.Class != OutlookDesktopCalendar.AppointmentClass)
        {
            return null;
        }

        return new OutlookAppointment
        {
            EntryId = item.EntryID as string ?? string.Empty,
            Subject = item.Subject as string,
            StartUtc = (DateTime)item.StartUTC,
            EndUtc = (DateTime)item.EndUTC,
            AllDayEvent = (bool)item.AllDayEvent,
            BusyStatus = (int)item.BusyStatus,
            ResponseStatus = (int)item.ResponseStatus,
            MeetingStatus = (int)item.MeetingStatus,
            IsRecurring = (bool)item.IsRecurring,
            Location = item.Location as string,
            TeamsMeetingUrl = ReadStringProperty(comItem, OutlookDesktopCalendar.TeamsMeetingUrlProperty),
        };
    }

    private static string? ReadStringProperty(object comItem, string schemaName)
    {
        object? accessor = null;
        try
        {
            dynamic item = comItem;
            accessor = item.PropertyAccessor;
            dynamic properties = accessor!;
            return properties.GetProperty(schemaName) as string;
        }
        catch (COMException)
        {
            // The property isn't set on this item.
            return null;
        }
        finally
        {
            Release(accessor);
        }
    }

    private static bool TryDisplayItem(ComScope com, dynamic session, OutlookOpenRequest request)
    {
        try
        {
            dynamic item = com.Track(session.GetItemFromID(request.EntryId));
            if (request.Start is { } start && (bool)item.IsRecurring)
            {
                try
                {
                    dynamic pattern = com.Track(item.GetRecurrencePattern());
                    dynamic occurrence = com.Track(pattern.GetOccurrence(start.ToLocalTime().DateTime));
                    occurrence.Display(false);
                    return true;
                }
                catch (COMException)
                {
                    // The occurrence was moved or deleted; fall back to the series.
                }
            }

            item.Display(false);
            return true;
        }
        catch (COMException ex)
        {
            Log.Error("Couldn't open the Outlook item; showing the calendar instead", ex);
            return false;
        }
    }

    private static object CreateApplication()
    {
        var type = Type.GetTypeFromProgID(ProgId) ?? throw new OutlookUnavailableException("Classic Outlook isn't installed.");

        // Outlook is single-instance, so this attaches to the running Outlook rather than starting another one.
        return Activator.CreateInstance(type) ?? throw new OutlookUnavailableException("Couldn't connect to Outlook.");
    }

    private static OutlookUnavailableException Translate(COMException ex) => (uint)ex.HResult switch
    {
        0x80010001 or 0x8001010A => new("Outlook is busy. The widget will try again shortly.", ex),
        0x80080005 => new("Couldn't connect to Outlook. If Outlook runs as administrator, restart it normally.", ex),
        0x800706BA or 0x800706BE or 0x80010108 => new("Outlook closed while the widget was reading your calendar.", ex),
        _ => new($"Outlook couldn't read your calendar ({ex.Message.Split('\n', 2)[0].Trim()})", ex),
    };

    private static bool IsRunning(string processName)
    {
        using var self = Process.GetCurrentProcess();
        var sessionId = self.SessionId;
        var processes = Process.GetProcessesByName(processName);
        try
        {
            return processes.Any(p => p.SessionId == sessionId);
        }
        finally
        {
            foreach (var process in processes)
            {
                process.Dispose();
            }
        }
    }

    private static void AllowOutlookToComeToFront()
    {
        var processes = Process.GetProcessesByName("OUTLOOK");
        foreach (var process in processes)
        {
            _ = AllowSetForegroundWindow(process.Id);
            process.Dispose();
        }
    }

    private static void Release(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
        {
            Marshal.ReleaseComObject(comObject);
        }
    }

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool AllowSetForegroundWindow(int processId);

    /// <summary>Releases every COM object handed out by Outlook, so Outlook doesn't keep items open.</summary>
    private sealed class ComScope : IDisposable
    {
        private readonly Stack<object?> objects = new();

        public object Track(object comObject)
        {
            objects.Push(comObject);
            return comObject;
        }

        public void Dispose()
        {
            while (objects.TryPop(out var comObject))
            {
                Release(comObject);
            }
        }
    }
}
