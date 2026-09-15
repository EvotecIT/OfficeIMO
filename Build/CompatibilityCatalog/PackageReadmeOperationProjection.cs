using System.Text;
using OfficeIMO;

internal static class PackageReadmeOperationProjection {
    private const string StartMarker = "<!-- officeimo-operation-catalog:start -->";
    private const string EndMarker = "<!-- officeimo-operation-catalog:end -->";
    private static readonly UTF8Encoding Utf8WithoutBom = new(encoderShouldEmitUTF8Identifier: false);

    internal static IReadOnlyDictionary<string, string> Create(
        string repositoryRoot,
        IReadOnlyList<OfficeOperationCapability> capabilities) {
        string root = Path.GetFullPath(repositoryRoot);
        var outputs = new SortedDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (IGrouping<string, OfficeOperationCapability> package in capabilities
                     .GroupBy(static row => row.PackageId, StringComparer.Ordinal)
                     .OrderBy(static group => group.Key, StringComparer.Ordinal)) {
            string? readmePath = ResolvePackageReadme(root, package.Key);
            if (readmePath == null) continue;
            string current = Read(root, readmePath);
            outputs[readmePath] = ReplaceBlock(current, Render(package.Key, package.ToArray()));
        }

        foreach (string readmePath in EnumerateProjectedReadmes(root)) {
            if (outputs.ContainsKey(readmePath)) continue;
            string current = Read(root, readmePath);
            outputs[readmePath] = ReplaceBlock(current, block: null);
        }
        return outputs;
    }

    internal static string Read(string repositoryRoot, string path) {
        string root = Path.GetFullPath(repositoryRoot);
        string validatedPath = ValidateProjectionPath(root, path);
        return Normalize(File.ReadAllText(validatedPath));
    }

    internal static void Write(string repositoryRoot, string path, string content) {
        string root = Path.GetFullPath(repositoryRoot);
        string validatedPath = ValidateProjectionPath(root, path);
        string directory = Path.GetDirectoryName(validatedPath)
            ?? throw new InvalidDataException("Package README path has no containing directory.");
        string temporaryPath = Path.Combine(directory, $".{Path.GetFileName(validatedPath)}.{Guid.NewGuid():N}.tmp");
        try {
            using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var writer = new StreamWriter(stream, Utf8WithoutBom)) {
                writer.Write(Normalize(content));
            }
            ValidateProjectionPath(root, validatedPath);
            File.Move(temporaryPath, validatedPath, overwrite: true);
        } finally {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }

    private static string? ResolvePackageReadme(string root, string packageId) {
        ValidatePackageId(packageId);
        string packageDirectory = ValidateProjectionPath(root, Path.Combine(root, packageId));
        if (!Directory.Exists(packageDirectory)) return null;
        return FindReadme(root, packageDirectory);
    }

    private static IEnumerable<string> EnumerateProjectedReadmes(string root) {
        string[] ignoredDirectories = { ".git", ".vs", "bin", "obj", "node_modules", "artifacts" };
        var pending = new Stack<string>();
        pending.Push(root);
        while (pending.Count > 0) {
            string directory = pending.Pop();
            string? readmePath = FindReadme(root, directory);
            if (readmePath != null) {
                string content = Read(root, readmePath);
                if (content.Contains(StartMarker, StringComparison.Ordinal) ||
                    content.Contains(EndMarker, StringComparison.Ordinal)) yield return readmePath;
            }
            foreach (string child in Directory.EnumerateDirectories(directory)) {
                if (ignoredDirectories.Contains(Path.GetFileName(child), StringComparer.OrdinalIgnoreCase)) continue;
                if ((File.GetAttributes(child) & FileAttributes.ReparsePoint) != 0) continue;
                pending.Push(ValidateProjectionPath(root, child));
            }
        }
    }

    private static string? FindReadme(string root, string directory) {
        string[] matches = Directory.EnumerateFiles(directory, "*", SearchOption.TopDirectoryOnly)
            .Where(static path => string.Equals(Path.GetFileName(path), "README.md", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (matches.Length > 1) {
            throw new InvalidDataException($"Directory contains multiple case-variant README files: {directory}");
        }
        return matches.Length == 0 ? null : ValidateProjectionPath(root, matches[0]);
    }

    private static void ValidatePackageId(string packageId) {
        if (string.IsNullOrWhiteSpace(packageId) || packageId is "." or ".." ||
            Path.IsPathRooted(packageId) ||
            packageId.IndexOfAny(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, '/', '\\' }) >= 0 ||
            packageId.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
            !string.Equals(packageId, Path.GetFileName(packageId), StringComparison.Ordinal)) {
            throw new InvalidDataException($"Package id must be a safe repository directory name: {packageId}");
        }
    }

    private static string ValidateProjectionPath(string root, string path) {
        string fullRoot = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string fullPath = Path.GetFullPath(path);
        string rootPrefix = fullRoot + Path.DirectorySeparatorChar;
        StringComparison comparison = OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        if (!fullPath.StartsWith(rootPrefix, comparison)) {
            throw new InvalidDataException($"Package README path escapes the repository root: {path}");
        }

        string relative = Path.GetRelativePath(fullRoot, fullPath);
        string current = fullRoot;
        foreach (string component in relative.Split(
                     new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
                     StringSplitOptions.RemoveEmptyEntries)) {
            current = Path.Combine(current, component);
            if (!File.Exists(current) && !Directory.Exists(current)) continue;
            if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0) {
                throw new InvalidDataException($"Package README path cannot traverse a filesystem link: {path}");
            }
        }
        return fullPath;
    }

    private static string Render(string packageId, IReadOnlyList<OfficeOperationCapability> rows) {
        var output = new StringBuilder();
        output.AppendLine(StartMarker)
            .AppendLine("## Generated capability summary")
            .AppendLine()
            .AppendLine("This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.")
            .AppendLine()
            .AppendLine("| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |")
            .AppendLine("| --- | ---: | ---: | ---: | ---: | ---: | ---: |");
        foreach (IGrouping<OfficeOperationKind, OfficeOperationCapability> operation in rows
                     .GroupBy(static row => row.Operation)
                     .OrderBy(static group => group.Key)) {
            output.Append("| ").Append(operation.Key);
            foreach (OfficeOperationSupportState state in Enum.GetValues<OfficeOperationSupportState>()) {
                output.Append(" | ").Append(operation.Count(row => row.State == state));
            }
            output.AppendLine(" |");
        }
        output.AppendLine()
            .Append("The complete rows for `").Append(packageId)
            .AppendLine("` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).")
            .AppendLine(EndMarker);
        return output.ToString();
    }

    private static string ReplaceBlock(string content, string? block) {
        ValidateMarkers(content);
        int start = content.IndexOf(StartMarker, StringComparison.Ordinal);
        int end = content.IndexOf(EndMarker, StringComparison.Ordinal);

        string withoutBlock = start >= 0
            ? JoinWithoutBlock(content, start, end + EndMarker.Length)
            : content.TrimEnd('\n');
        if (block == null) return withoutBlock.Length == 0 ? string.Empty : withoutBlock + "\n";
        return withoutBlock + "\n\n" + block.TrimEnd('\n') + "\n";
    }

    private static string JoinWithoutBlock(string content, int start, int after) {
        string before = content.Substring(0, start).TrimEnd('\n');
        string trailing = content.Substring(after).Trim('\n');
        if (before.Length == 0) return trailing;
        if (trailing.Length == 0) return before;
        return before + "\n\n" + trailing;
    }

    private static void ValidateMarkers(string content) {
        string[] lines = content.Split('\n');
        int start = -1;
        int end = -1;
        int startCount = 0;
        int endCount = 0;
        for (int index = 0; index < lines.Length; index++) {
            string line = lines[index];
            bool containsStart = line.Contains(StartMarker, StringComparison.Ordinal);
            bool containsEnd = line.Contains(EndMarker, StringComparison.Ordinal);
            if ((containsStart && line != StartMarker) || (containsEnd && line != EndMarker)) {
                throw new InvalidDataException("Package README generated operation markers must appear on standalone lines.");
            }
            if (containsStart) {
                start = index;
                startCount++;
            }
            if (containsEnd) {
                end = index;
                endCount++;
            }
        }
        if (startCount == 0 && endCount == 0) return;
        if (startCount != 1 || endCount != 1 || start >= end) {
            throw new InvalidDataException("Package README must contain exactly one correctly ordered generated operation block.");
        }
    }

    private static string Normalize(string value) => value
        .Replace("\r\n", "\n")
        .Replace("\r", "\n");
}
