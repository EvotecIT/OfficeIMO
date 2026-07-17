using System.Runtime.InteropServices;

namespace OfficeIMO.Email.Store;

internal static partial class EmailStorePathIdentity {
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

    internal static bool IsSameOrDescendant(string candidatePath, string rootPath) {
        string candidate = TrimEndingDirectorySeparators(Path.GetFullPath(candidatePath));
        string root = TrimEndingDirectorySeparators(Path.GetFullPath(rootPath));
        bool candidateResolved = TryResolvePhysicalPath(candidate, out string resolvedCandidate);
        bool rootResolved = TryResolvePhysicalPath(root, out string resolvedRoot);
        if ((!candidateResolved && ContainsReparsePointInExistingAncestry(candidate)) ||
            (!rootResolved && ContainsReparsePointInExistingAncestry(root))) {
            return true;
        }
        if (candidateResolved) candidate = resolvedCandidate;
        if (rootResolved) root = resolvedRoot;
        StringComparison comparison = GetComparison(root);
        if (string.Equals(candidate, root, comparison)) return true;
        string prefix = string.Concat(root, Path.DirectorySeparatorChar.ToString());
        return candidate.StartsWith(prefix, comparison);
    }

    internal static bool IsCaseInsensitiveFileSystem(string path) {
        string fullPath = Path.GetFullPath(path);
        string? existingPath = TrimEndingDirectorySeparators(fullPath);
        while (!string.IsNullOrEmpty(existingPath) &&
               !File.Exists(existingPath) && !Directory.Exists(existingPath)) {
            existingPath = Path.GetDirectoryName(existingPath);
        }
        if (string.IsNullOrEmpty(existingPath)) {
            return RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
        }

        string? directory = Directory.Exists(existingPath)
            ? existingPath
            : Path.GetDirectoryName(existingPath);
        if (string.IsNullOrEmpty(directory)) {
            return RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) &&
            TryGetWindowsDirectoryCaseInsensitive(directory, out bool caseInsensitive)) {
            return caseInsensitive;
        }
        if (Directory.Exists(directory)) {
            try {
                foreach (string entry in Directory.EnumerateFileSystemEntries(directory)) {
                    if (TryDetectFromExistingName(entry, out caseInsensitive)) return caseInsensitive;
                }
            } catch (IOException) {
            } catch (UnauthorizedAccessException) {
            }
        }
        return RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
    }

    private static bool TryDetectFromExistingName(string path, out bool caseInsensitive) {
        caseInsensitive = false;
        string existingPath = TrimEndingDirectorySeparators(path);
        string? parent = Path.GetDirectoryName(existingPath);
        string name = Path.GetFileName(existingPath);
        if (string.IsNullOrEmpty(parent) || string.IsNullOrEmpty(name)) return false;

        int letterIndex = -1;
        for (int index = name.Length - 1; index >= 0; index--) {
            if (char.IsLetter(name[index])) {
                letterIndex = index;
                break;
            }
        }
        if (letterIndex < 0) return false;

        char original = name[letterIndex];
        char alternate = char.IsUpper(original) ? char.ToLowerInvariant(original) : char.ToUpperInvariant(original);
        if (alternate == original) return false;
        var alternateName = new StringBuilder(name);
        alternateName[letterIndex] = alternate;
        string alternatePath = Path.Combine(parent, alternateName.ToString());
        if (!File.Exists(alternatePath) && !Directory.Exists(alternatePath)) {
            caseInsensitive = false;
            return true;
        }

        try {
            int matches = Directory.EnumerateFileSystemEntries(parent)
                .Count(entry => string.Equals(Path.GetFileName(entry), name,
                    StringComparison.OrdinalIgnoreCase));
            caseInsensitive = matches <= 1;
            return true;
        } catch (IOException) {
            return false;
        } catch (UnauthorizedAccessException) {
            return false;
        }
    }

    private static bool ContainsReparsePointInExistingAncestry(string path) {
        string? candidate = TrimEndingDirectorySeparators(Path.GetFullPath(path));
        while (!string.IsNullOrEmpty(candidate)) {
            try {
                if ((File.Exists(candidate) || Directory.Exists(candidate)) &&
                    (File.GetAttributes(candidate) & FileAttributes.ReparsePoint) != 0) {
                    return true;
                }
            } catch (IOException) {
            } catch (UnauthorizedAccessException) {
            }

            string? parent = Path.GetDirectoryName(candidate);
            if (string.IsNullOrEmpty(parent) || string.Equals(parent, candidate, StringComparison.Ordinal)) break;
            candidate = parent;
        }
        return false;
    }

    private static string TrimEndingDirectorySeparators(string path) {
        string root = Path.GetPathRoot(path) ?? string.Empty;
        int length = path.Length;
        while (length > root.Length &&
               (path[length - 1] == Path.DirectorySeparatorChar ||
                path[length - 1] == Path.AltDirectorySeparatorChar)) {
            length--;
        }
        return length == path.Length ? path : path.Substring(0, length);
    }
}
