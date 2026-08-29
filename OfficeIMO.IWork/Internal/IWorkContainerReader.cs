using System.IO.Compression;

namespace OfficeIMO.IWork.Internal;

internal sealed class IWorkPackageData {
    internal IWorkPackageData(IWorkContainerKind containerKind, IReadOnlyList<IWorkPackageEntry> entries) {
        ContainerKind = containerKind;
        Entries = entries;
    }

    internal IWorkContainerKind ContainerKind { get; }
    internal IReadOnlyList<IWorkPackageEntry> Entries { get; }
}

internal static class IWorkContainerReader {
    internal static IWorkPackageData Read(string path, IWorkReadOptions options) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A source path is required.", nameof(path));
        if (Directory.Exists(path)) return ReadDirectory(path, options);
        if (!File.Exists(path)) throw new FileNotFoundException("The iWork source was not found.", path);

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            bufferSize: 81920, FileOptions.SequentialScan);
        return Read(stream, options);
    }

    internal static IWorkPackageData Read(Stream stream, IWorkReadOptions options) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The source stream must be readable.", nameof(stream));
        byte[] package = ReadBounded(stream, options.MaximumPackageBytes, "package");
        using var copy = new MemoryStream(package, writable: false);
        return ReadZip(copy, options);
    }

    private static IWorkPackageData ReadDirectory(string path, IWorkReadOptions options) {
        if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) {
            throw new InvalidDataException("Directory bundles cannot be symbolic links or reparse points.");
        }
        string root = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        var entries = new Dictionary<string, IWorkPackageEntry>(StringComparer.Ordinal);
        long total = 0;
        int nodeCount = 0;
        var directories = new Stack<string>();
        directories.Push(Path.GetFullPath(path));
        StringComparison pathComparison = Path.DirectorySeparatorChar == '\\'
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        while (directories.Count > 0) {
            string directory = directories.Pop();
            foreach (string fileSystemEntry in Directory.EnumerateFileSystemEntries(directory, "*", SearchOption.TopDirectoryOnly)) {
                EnforceEntryCount(ref nodeCount, options);
                FileAttributes attributes = File.GetAttributes(fileSystemEntry);
                if ((attributes & FileAttributes.ReparsePoint) != 0) {
                    throw new InvalidDataException($"Directory bundles cannot contain symbolic-link entries: {fileSystemEntry}.");
                }
                if ((attributes & FileAttributes.Directory) != 0) {
                    string directoryFullPath = Path.GetFullPath(fileSystemEntry);
                    if (!directoryFullPath.StartsWith(root, pathComparison)) {
                        throw new InvalidDataException("A bundle entry resolves outside the source directory.");
                    }
                    _ = NormalizePath(directoryFullPath.Substring(root.Length));
                    directories.Push(fileSystemEntry);
                    continue;
                }
                string full = Path.GetFullPath(fileSystemEntry);
                if (!full.StartsWith(root, pathComparison)) throw new InvalidDataException("A bundle entry resolves outside the source directory.");
                string relative = NormalizePath(full.Substring(root.Length));
                long remainingPackageBytes = options.MaximumPackageBytes - total;
                long remainingEntryBytes = options.MaximumTotalEntryBytes - total;
                long readLimit = Math.Min(options.MaximumEntryBytes,
                    Math.Min(remainingPackageBytes, remainingEntryBytes));
                if (readLimit < 0) {
                    throw new InvalidDataException("Directory bundle size exceeds a configured package limit.");
                }
                byte[] bytes;
                using (var input = new FileStream(full, FileMode.Open, FileAccess.Read, FileShare.Read,
                           bufferSize: 81920, FileOptions.SequentialScan)) {
                    bytes = ReadBounded(input, readLimit, relative);
                }
                EnforceEntryBounds(bytes.LongLength, ref total, options, relative);
                AddEntry(entries, relative, bytes);
            }
        }
        ExpandNestedIndex(entries, ref total, ref nodeCount, options);
        IWorkContainerKind kind = entries.ContainsKey("Index.zip")
            ? IWorkContainerKind.ZipPackageWithNestedIndex
            : IWorkContainerKind.DirectoryBundle;
        return new IWorkPackageData(kind, entries.Values.OrderBy(entry => entry.Path, StringComparer.Ordinal).ToArray());
    }

    private static IWorkPackageData ReadZip(Stream stream, IWorkReadOptions options) {
        var entries = new Dictionary<string, IWorkPackageEntry>(StringComparer.Ordinal);
        long total = 0;
        int nodeCount = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true)) {
            ReadArchiveEntries(archive, entries, prefix: null, ref total, ref nodeCount, options);
        }
        bool nested = entries.ContainsKey("Index.zip");
        ExpandNestedIndex(entries, ref total, ref nodeCount, options);
        return new IWorkPackageData(
            nested ? IWorkContainerKind.ZipPackageWithNestedIndex : IWorkContainerKind.ZipPackage,
            entries.Values.OrderBy(entry => entry.Path, StringComparer.Ordinal).ToArray());
    }

    private static void ExpandNestedIndex(Dictionary<string, IWorkPackageEntry> entries, ref long total,
        ref int nodeCount, IWorkReadOptions options) {
        if (!entries.TryGetValue("Index.zip", out IWorkPackageEntry? nested)) return;
        using var stream = new MemoryStream(nested.Bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        ReadArchiveEntries(archive, entries, "Index", ref total, ref nodeCount, options);
    }

    private static void ReadArchiveEntries(ZipArchive archive, Dictionary<string, IWorkPackageEntry> entries,
        string? prefix, ref long total, ref int nodeCount, IWorkReadOptions options) {
        foreach (ZipArchiveEntry entry in archive.Entries) {
            EnforceEntryCount(ref nodeCount, options);
            if (string.IsNullOrEmpty(entry.Name)) {
                string directoryPath = entry.FullName.TrimEnd('/', '\\');
                if (directoryPath.Length > 0) _ = NormalizePath(directoryPath);
                continue;
            }
            string normalized = NormalizePath(entry.FullName);
            if (!string.IsNullOrEmpty(prefix) && !normalized.StartsWith(prefix + "/", StringComparison.Ordinal)) {
                normalized = prefix + "/" + normalized;
            }
            EnforceEntryBounds(entry.Length, ref total, options, normalized);
            using Stream input = entry.Open();
            byte[] bytes = ReadBounded(input, Math.Min(options.MaximumEntryBytes, entry.Length), normalized);
            if (bytes.LongLength != entry.Length) throw new InvalidDataException($"Entry {normalized} changed length while it was read.");
            AddEntry(entries, normalized, bytes);
        }
    }

    private static void EnforceEntryCount(ref int count, IWorkReadOptions options) {
        if (count >= options.MaximumEntryCount) {
            throw new InvalidDataException($"Package entry count exceeds the configured limit of {options.MaximumEntryCount}.");
        }
        count++;
    }

    private static void EnforceEntryBounds(long length, ref long total, IWorkReadOptions options, string path) {
        if (length < 0 || length > options.MaximumEntryBytes) {
            throw new InvalidDataException($"Entry {path} has length {length}, above the configured limit of {options.MaximumEntryBytes} bytes.");
        }
        if (total > options.MaximumTotalEntryBytes - length) {
            throw new InvalidDataException($"Combined package entries exceed the configured limit of {options.MaximumTotalEntryBytes} bytes.");
        }
        total += length;
    }

    private static byte[] ReadBounded(Stream stream, long maximumBytes, string label) {
        using var output = new MemoryStream();
        var buffer = new byte[81920];
        while (true) {
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            if (output.Length > maximumBytes - read) {
                throw new InvalidDataException($"The {label} exceeds the configured limit of {maximumBytes} bytes.");
            }
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

    private static string NormalizePath(string path) {
        if (string.IsNullOrEmpty(path) || path[0] == '/' || path[0] == '\\' || Path.IsPathRooted(path)) {
            throw new InvalidDataException($"Package entry paths must be relative: {path}.");
        }
        string normalized = path.Replace('\\', '/').TrimStart('/');
        string[] segments = normalized.Split('/');
        if (segments.Length == 0 || segments.Any(segment => segment.Length == 0 || segment == "." || segment == ".." || segment.Contains(':'))) {
            throw new InvalidDataException($"Unsafe or empty package entry path: {path}.");
        }
        return string.Join("/", segments);
    }

    private static void AddEntry(Dictionary<string, IWorkPackageEntry> entries, string path, byte[] bytes) {
        if (entries.ContainsKey(path)) throw new InvalidDataException($"Duplicate package entry path: {path}.");
        entries.Add(path, new IWorkPackageEntry(path, bytes));
    }
}
