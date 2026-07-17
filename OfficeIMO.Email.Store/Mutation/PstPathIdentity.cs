using System.Runtime.InteropServices;

namespace OfficeIMO.Email.Store;

internal static class PstPathIdentity {
    internal static bool IsCaseInsensitivePlatform =>
        RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ||
        RuntimeInformation.IsOSPlatform(OSPlatform.OSX);

    internal static StringComparison Comparison => IsCaseInsensitivePlatform
        ? StringComparison.OrdinalIgnoreCase
        : StringComparison.Ordinal;

    internal static string Normalize(string path) => Normalize(path, IsCaseInsensitivePlatform);

    internal static string Normalize(string path, bool caseInsensitive) {
        string identity = Path.GetFullPath(path);
        return caseInsensitive ? identity.ToUpperInvariant() : identity;
    }
}
