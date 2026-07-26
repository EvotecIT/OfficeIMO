namespace OfficeIMO.Tool;

/// <summary>Provides link-aware path comparisons shared by command areas that write files.</summary>
internal static class OfficeImoToolPathSafety {
    /// <summary>Determines whether two paths resolve to the same file-system location.</summary>
    internal static bool PathsEqual(string firstPath, string secondPath) =>
        string.Equals(
            ResolveExistingLinks(firstPath),
            ResolveExistingLinks(secondPath),
            PathComparison);

    /// <summary>Resolves every existing symbolic-link segment while preserving a non-existing tail.</summary>
    internal static string ResolveExistingLinks(string path) {
        string current = Path.GetFullPath(path);
        for (int pass = 0; pass < 40; pass++) {
            string resolved = ResolveLinkPass(current);
            if (string.Equals(current, resolved, PathComparison)) return resolved;
            current = resolved;
        }

        throw new IOException("Linked path resolution exceeded the supported depth.");
    }

    /// <summary>Determines whether a resolved candidate is the parent path itself or one of its descendants.</summary>
    internal static bool IsSameOrChildPath(string parentPath, string candidatePath) {
        if (string.Equals(parentPath, candidatePath, PathComparison)) return true;
        string parentPrefix = Path.EndsInDirectorySeparator(parentPath)
            ? parentPath
            : parentPath + Path.DirectorySeparatorChar;
        return candidatePath.StartsWith(parentPrefix, PathComparison);
    }

    private static string ResolveLinkPass(string path) {
        string fullPath = Path.GetFullPath(path);
        string? root = Path.GetPathRoot(fullPath);
        if (string.IsNullOrEmpty(root)) return fullPath;

        string current = root;
        string relative = fullPath.Substring(root.Length);
        string[] segments = relative.Split(
            new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
            StringSplitOptions.RemoveEmptyEntries);
        for (int index = 0; index < segments.Length; index++) {
            current = Path.Combine(current, segments[index]);
            bool isDirectory = Directory.Exists(current);
            bool isFile = File.Exists(current);
            if (!isDirectory && !isFile) {
                for (int remainder = index + 1; remainder < segments.Length; remainder++) {
                    current = Path.Combine(current, segments[remainder]);
                }
                break;
            }

            FileSystemInfo link = isDirectory
                ? new DirectoryInfo(current)
                : new FileInfo(current);
            try {
                if (link.LinkTarget == null) continue;
                FileSystemInfo? target = link.ResolveLinkTarget(returnFinalTarget: true);
                if (target == null) {
                    throw new IOException("Could not resolve linked path '" + current + "'.");
                }
                current = Path.GetFullPath(target.FullName);
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                throw new IOException("Could not resolve linked path '" + current + "'.", exception);
            }
        }

        return Path.TrimEndingDirectorySeparator(Path.GetFullPath(current));
    }

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
}
