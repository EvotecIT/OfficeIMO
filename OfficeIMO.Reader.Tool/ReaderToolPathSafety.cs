namespace OfficeIMO.Reader.Tool;

internal static class ReaderToolPathSafety {
    internal static void EnsureOutsideInput(string inputPath, params string?[] candidatePaths) {
        string resolvedInput = ResolveExistingLinks(inputPath);
        foreach (string? candidatePath in candidatePaths) {
            if (string.IsNullOrWhiteSpace(candidatePath)) continue;
            string resolvedCandidate = ResolveExistingLinks(candidatePath!);
            if (IsSameOrChildPath(resolvedInput, resolvedCandidate)) {
                throw new ReaderToolOutputException(
                    "Folder output and asset directories must be outside the input folder.");
            }
        }
    }

    private static string ResolveExistingLinks(string path) {
        string current = Path.GetFullPath(path);
        StringComparison comparison = OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        for (int pass = 0; pass < 40; pass++) {
            string resolved = ResolveLinkPass(current);
            if (string.Equals(current, resolved, comparison)) return resolved;
            current = resolved;
        }

        throw new ReaderToolOutputException("Linked path resolution exceeded the supported depth.");
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
                    throw new ReaderToolOutputException(
                        "Could not resolve linked path '" + current + "'.");
                }
                current = Path.GetFullPath(target.FullName);
            } catch (ReaderToolOutputException) {
                throw;
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                throw new ReaderToolOutputException(
                    "Could not resolve linked path '" + current + "'.",
                    exception);
            }
        }

        return Path.TrimEndingDirectorySeparator(Path.GetFullPath(current));
    }

    private static bool IsSameOrChildPath(string parentPath, string candidatePath) {
        StringComparison comparison = OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        if (string.Equals(parentPath, candidatePath, comparison)) return true;
        string parentPrefix = Path.EndsInDirectorySeparator(parentPath)
            ? parentPath
            : parentPath + Path.DirectorySeparatorChar;
        return candidatePath.StartsWith(parentPrefix, comparison);
    }
}
