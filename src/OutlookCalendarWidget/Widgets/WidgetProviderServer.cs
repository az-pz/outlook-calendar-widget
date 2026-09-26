using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using OutlookCalendarWidget.Services;

namespace OutlookCalendarWidget.Widgets;

/// <summary>Registers the widget provider class object with COM and keeps the process alive while widgets exist.</summary>
internal static partial class WidgetProviderServer
{
    /// <summary>Must match the CLSID declared for the ComServer and WidgetProvider extensions in Package.appxmanifest.</summary>
    public static readonly Guid ProviderClsid = new("E5BF83A9-075E-4D20-80B4-B4EF2B16017C");

    private const uint ClsctxLocalServer = 0x4;
    private const uint RegclsMultipleUse = 0x1;
    private const uint CoinitMultithreaded = 0x0;

    public static void Run()
    {
        // Register from a dedicated MTA thread so incoming widget calls are serviced by the COM thread pool
        // rather than needing a message pump on the (STA) main thread.
        var thread = new Thread(Serve) { Name = "WidgetProviderServer", IsBackground = false };
        thread.SetApartmentState(ApartmentState.MTA);
        thread.Start();
        thread.Join();
    }

    private static void Serve()
    {
        _ = CoInitializeEx(0, CoinitMultithreaded);

        var factory = new WidgetProviderFactory();
        var comWrappers = new StrategyBasedComWrappers();
        var factoryPointer = comWrappers.GetOrCreateComInterfaceForObject(factory, CreateComInterfaceFlags.None);
        try
        {
            Marshal.ThrowExceptionForHR(CoRegisterClassObject(ProviderClsid, factoryPointer, ClsctxLocalServer, RegclsMultipleUse, out var cookie));
            Log.Info("Widget provider class object registered.");
            try
            {
                CalendarWidgetProvider.NoWidgetsRemaining.WaitOne();
                Log.Info("No widgets remaining; widget provider exiting.");
            }
            finally
            {
                _ = CoRevokeClassObject(cookie);
            }
        }
        finally
        {
            Marshal.Release(factoryPointer);
            GC.KeepAlive(factory);
        }
    }

    [LibraryImport("ole32.dll")]
    private static partial int CoInitializeEx(nint reserved, uint coInit);

    [LibraryImport("ole32.dll")]
    private static partial int CoRegisterClassObject(in Guid rclsid, nint pUnk, uint dwClsContext, uint flags, out uint lpdwRegister);

    [LibraryImport("ole32.dll")]
    private static partial int CoRevokeClassObject(uint dwRegister);
}
