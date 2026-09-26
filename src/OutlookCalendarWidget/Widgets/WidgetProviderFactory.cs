using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Microsoft.Windows.Widgets.Providers;
using OutlookCalendarWidget.Services;
using WinRT;

namespace OutlookCalendarWidget.Widgets;

/// <summary>Classic COM <c>IClassFactory</c>, declared with source-generated COM interop (trimming/AOT friendly).</summary>
/// <remarks><c>riid</c> must be <c>in</c>: callers pass IIDs from read-only memory, and a <c>ref</c> parameter would be written back.</remarks>
[GeneratedComInterface]
[Guid("00000001-0000-0000-C000-000000000046")]
internal partial interface IClassFactory
{
    [PreserveSig]
    int CreateInstance(nint pUnkOuter, in Guid riid, out nint ppvObject);

    [PreserveSig]
    int LockServer([MarshalAs(UnmanagedType.Bool)] bool fLock);
}

/// <summary>Hands the Widgets host a WinRT object implementing <see cref="IWidgetProvider"/>.</summary>
[GeneratedComClass]
internal sealed partial class WidgetProviderFactory : IClassFactory
{
    private const int ClassENoAggregation = unchecked((int)0x80040110);

    public int CreateInstance(nint pUnkOuter, in Guid riid, out nint ppvObject)
    {
        ppvObject = 0;
        if (pUnkOuter != 0)
        {
            return ClassENoAggregation;
        }

        try
        {
            var unknown = MarshalInspectable<IWidgetProvider>.FromManaged(CalendarWidgetProvider.Instance);
            try
            {
                return Marshal.QueryInterface(unknown, in riid, out ppvObject);
            }
            finally
            {
                Marshal.Release(unknown);
            }
        }
        catch (Exception ex)
        {
            Log.Error("Failed to create the widget provider", ex);
            return ex.HResult;
        }
    }

    public int LockServer(bool fLock) => 0;
}
