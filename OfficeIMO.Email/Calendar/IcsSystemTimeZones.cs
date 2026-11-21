#if NETFRAMEWORK
using System.Runtime.InteropServices;
using System.Text;
#endif

namespace OfficeIMO.Email;

/// <summary>Finds the system time zone named by an iCalendar TZID, as an IANA or a Windows identifier.</summary>
internal static class IcsSystemTimeZones {
    /// <summary>
    /// Resolves <paramref name="id"/> to a system zone. .NET 8 and later accept IANA and Windows identifiers on every
    /// operating system; .NET Framework knows only Windows identifiers, so an IANA identifier is first mapped through
    /// the ICU library that ships with Windows, which is the same mapping .NET 8 uses there.
    /// </summary>
    internal static bool TryFind(string id, out TimeZoneInfo zone) {
        zone = null!;
        if (string.IsNullOrWhiteSpace(id)) return false;
        try {
            zone = TimeZoneInfo.FindSystemTimeZoneById(id);
            return true;
        } catch (TimeZoneNotFoundException) {
        } catch (InvalidTimeZoneException) {
            return false;
        }
#if NETFRAMEWORK
        string? windowsId = IanaToWindowsId(id);
        if (windowsId == null || string.Equals(windowsId, id, StringComparison.OrdinalIgnoreCase)) return false;
        try {
            zone = TimeZoneInfo.FindSystemTimeZoneById(windowsId);
            return true;
        } catch (TimeZoneNotFoundException) {
        } catch (InvalidTimeZoneException) {
        }
#endif
        return false;
    }

#if NETFRAMEWORK
    private static int _icuUnavailable;

    private static string? IanaToWindowsId(string id) {
        if (Volatile.Read(ref _icuUnavailable) != 0) return null;
        try {
            var buffer = new StringBuilder(128);
            int status = 0;
            int length = GetWindowsTimeZoneId(id, id.Length, buffer, buffer.Capacity, ref status);
            // ICU reports failures as positive status codes and an unknown zone as an empty result.
            return status > 0 || length <= 0 || length > buffer.Capacity ? null : buffer.ToString(0, length);
        } catch (Exception exception) when (exception is DllNotFoundException || exception is EntryPointNotFoundException) {
            // Windows versions before 10 version 1903 ship no icu.dll; IANA identifiers stay unresolved there.
            Volatile.Write(ref _icuUnavailable, 1);
            return null;
        }
    }

    [DllImport("icu.dll", EntryPoint = "ucal_getWindowsTimeZoneID", CharSet = CharSet.Unicode, ExactSpelling = true,
        CallingConvention = CallingConvention.Cdecl)]
    private static extern int GetWindowsTimeZoneId(string id, int length, StringBuilder windowsId, int capacity, ref int status);
#endif
}
