using System.Runtime.InteropServices;

namespace OfficeIMO.Email.Store;

internal static class EmailStorePathIdentity {
    internal static string Normalize(string path) =>
        Normalize(path, IsCaseInsensitiveFileSystem(path));

    internal static bool AreEquivalent(string left, string right) {
        string leftPath = Path.GetFullPath(left);
        string rightPath = Path.GetFullPath(right);
        if (string.Equals(leftPath, rightPath, StringComparison.Ordinal)) return true;
        return IsCaseInsensitiveFileSystem(leftPath) &&
            IsCaseInsensitiveFileSystem(rightPath) &&
            string.Equals(leftPath, rightPath, StringComparison.OrdinalIgnoreCase);
    }

    internal static string Normalize(string path, bool caseInsensitive) {
        string identity = Path.GetFullPath(path);
        return caseInsensitive ? identity.ToUpperInvariant() : identity;
    }

    internal static StringComparer GetComparer(string path) =>
        IsCaseInsensitiveFileSystem(path) ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;

    internal static StringComparison GetComparison(string path) =>
        IsCaseInsensitiveFileSystem(path) ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

    internal static bool IsCaseInsensitiveFileSystem(string path) {
        string fullPath = Path.GetFullPath(path);
        string? directory = File.Exists(fullPath) ? Path.GetDirectoryName(fullPath) : fullPath;
        while (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
            directory = Path.GetDirectoryName(directory);
        }
        if (string.IsNullOrEmpty(directory)) {
            return RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
        }

        string probeName = string.Concat(".officeimo-case-probe-",
            Guid.NewGuid().ToString("N"), "-a.tmp");
        string probePath = Path.Combine(directory, probeName);
        string alternatePath = Path.Combine(directory, probeName.ToUpperInvariant());
        try {
            using (new FileStream(probePath, FileMode.CreateNew, FileAccess.ReadWrite,
                       FileShare.ReadWrite | FileShare.Delete, 1, FileOptions.RandomAccess)) {
                return File.Exists(alternatePath);
            }
        } catch (IOException) {
            return RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
        } catch (UnauthorizedAccessException) {
            return RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
        } finally {
            try { if (File.Exists(probePath)) File.Delete(probePath); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
