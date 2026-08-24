using OfficeIMO.Internal;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Email.Store;

internal static class EmailStorePathIdentity {
    internal static string Normalize(string path) => OfficePathIdentity.Normalize(path);
    internal static string ResolvePhysicalPath(string path) => OfficePathIdentity.ResolvePhysicalPath(path);
    internal static string GetPhysicalIdentityKey(string path) => OfficePathIdentity.GetPhysicalIdentityKey(path);
    internal static string GetPhysicalIdentityKey(string path, SafeFileHandle handle) =>
        OfficePathIdentity.GetPhysicalIdentityKey(path, handle);
    internal static bool AreEquivalent(string left, string right) => OfficePathIdentity.AreEquivalent(left, right);
    internal static string Normalize(string path, bool caseInsensitive) => OfficePathIdentity.Normalize(path, caseInsensitive);
    internal static StringComparer GetComparer(string path) => OfficePathIdentity.GetComparer(path);
    internal static StringComparison GetComparison(string path) => OfficePathIdentity.GetComparison(path);
    internal static bool IsSameOrDescendant(string candidatePath, string rootPath) =>
        OfficePathIdentity.IsSameOrDescendant(candidatePath, rootPath);
    internal static bool IsCaseInsensitiveFileSystem(string path) =>
        OfficePathIdentity.IsCaseInsensitiveFileSystem(path);
}
