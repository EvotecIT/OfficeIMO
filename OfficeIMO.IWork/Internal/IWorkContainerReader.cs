using System.IO.Compression;
using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.IWork.Internal;

internal sealed class IWorkPackageData {
    internal IWorkPackageData(IWorkContainerKind containerKind, IReadOnlyList<IWorkPackageEntry> entries,
        long containerLengthBytes, IReadOnlyList<string>? sourceFilePaths = null) {
        ContainerKind = containerKind;
        Entries = entries;
        ContainerLengthBytes = containerLengthBytes;
        SourceFilePaths = sourceFilePaths ?? Array.Empty<string>();
    }

    internal IWorkContainerKind ContainerKind { get; }
    internal IReadOnlyList<IWorkPackageEntry> Entries { get; }
    internal long ContainerLengthBytes { get; }
    // Native paths cannot be reconstructed from normalized archive names on POSIX.
    internal IReadOnlyList<string> SourceFilePaths { get; }
}

internal static class IWorkContainerReader {
    internal static IWorkPackageData Read(string path, IWorkReadOptions options,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A source path is required.", nameof(path));
        if (Directory.Exists(path)) return ReadDirectory(path, options, cancellationToken);
        using FileStream stream = OpenPackageFileForRead(path);
        return Read(stream, options, cancellationToken);
    }

    internal static FileStream OpenPackageFileForRead(string path) => OpenPackageFileForRead(path, FileShare.Read);

    internal static FileStream OpenPackageFileForContentDetection(string path) =>
        OpenPackageFileForRead(path, FileShare.ReadWrite | FileShare.Delete);

    private static FileStream OpenPackageFileForRead(string path, FileShare share) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A source path is required.", nameof(path));
        if (!File.Exists(path)) throw new FileNotFoundException("The iWork source was not found.", path);

        string fullPath = Path.GetFullPath(path);
        if ((File.GetAttributes(fullPath) & FileAttributes.ReparsePoint) != 0) {
            throw new InvalidDataException("iWork package paths cannot be symbolic links or reparse points.");
        }
        string parent = Path.GetDirectoryName(fullPath)
            ?? throw new InvalidDataException("The iWork package path has no parent directory.");
        string physicalRoot = OfficePathIdentity.ResolvePhysicalPath(parent);
        return OfficePathIdentity.OpenRegularFileForRead(fullPath, physicalRoot, 81920, share);
    }

    internal static IWorkPackageData Read(Stream stream, IWorkReadOptions options,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The source stream must be readable.", nameof(stream));
        byte[] package = ReadBounded(stream, options.MaximumPackageBytes, "package", cancellationToken);
        using var copy = new MemoryStream(package, writable: false);
        return ReadZip(copy, options, cancellationToken);
    }

    internal static IWorkPackageData ReadDirectorySnapshot(string path, IWorkReadOptions options, CancellationToken cancellationToken) =>
        ReadDirectory(path, options, cancellationToken, expandNestedIndex: false);

    private static IWorkPackageData ReadDirectory(string path, IWorkReadOptions options,
        CancellationToken cancellationToken, bool expandNestedIndex = true) {
        using var rootHandle = OfficePathIdentity.OpenDirectoryForIdentity(path,
            out string physicalRoot);
        string root = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        var entries = new Dictionary<string, IWorkPackageEntry>(StringComparer.Ordinal);
        var sourceFilePaths = new List<string>();
        long total = 0;
        int nodeCount = 0;
        var directories = new Stack<string>();
        directories.Push(Path.GetFullPath(path));
        StringComparison pathComparison = Path.DirectorySeparatorChar == '\\'
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        while (directories.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            string directory = directories.Pop();
            OfficePathIdentity.EnsurePathMatchesOpenedDirectory(path, rootHandle);
            foreach (string fileSystemEntry in Directory.EnumerateFileSystemEntries(
                         directory, "*", SearchOption.TopDirectoryOnly)) {
                cancellationToken.ThrowIfCancellationRequested();
                EnforceEntryCount(ref nodeCount, options);
                OfficePathIdentity.EnsurePathMatchesOpenedDirectory(path, rootHandle);
                FileAttributes attributes = File.GetAttributes(fileSystemEntry);
                OfficePathIdentity.EnsurePathMatchesOpenedDirectory(path, rootHandle);
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
                using (FileStream input = OfficePathIdentity.OpenRegularFileForRead(full, physicalRoot, 81920)) {
                    bytes = ReadBounded(input, readLimit, relative, cancellationToken);
                }
                OfficePathIdentity.EnsurePathMatchesOpenedDirectory(path, rootHandle);
                EnforceEntryBounds(bytes.LongLength, ref total, options, relative);
                AddEntry(entries, relative, bytes);
                sourceFilePaths.Add(full);
            }
            OfficePathIdentity.EnsurePathMatchesOpenedDirectory(path, rootHandle);
        }
        OfficePathIdentity.EnsurePathMatchesOpenedDirectory(path, rootHandle);
        long containerLengthBytes = total;
        if (expandNestedIndex) ExpandNestedIndex(entries, ref total, ref nodeCount, options, cancellationToken);
        return new IWorkPackageData(IWorkContainerKind.DirectoryBundle,
            entries.Values.OrderBy(entry => entry.Path, StringComparer.Ordinal).ToArray(), containerLengthBytes,
            sourceFilePaths.ToArray());
    }

    private static IWorkPackageData ReadZip(Stream stream, IWorkReadOptions options,
        CancellationToken cancellationToken) {
        var entries = new Dictionary<string, IWorkPackageEntry>(StringComparer.Ordinal);
        long total = 0;
        int nodeCount = 0;
        ValidateZipCentralDirectory(stream, options.MaximumEntryCount - nodeCount,
            options.MaximumEntryCount, cancellationToken);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true)) {
            ReadArchiveEntries(archive, entries, prefix: null, ref total, ref nodeCount, options,
                cancellationToken);
        }
        bool nested = entries.ContainsKey("Index.zip");
        ExpandNestedIndex(entries, ref total, ref nodeCount, options, cancellationToken);
        return new IWorkPackageData(
            nested ? IWorkContainerKind.ZipPackageWithNestedIndex : IWorkContainerKind.ZipPackage,
            entries.Values.OrderBy(entry => entry.Path, StringComparer.Ordinal).ToArray(), stream.Length);
    }

    private static void ExpandNestedIndex(Dictionary<string, IWorkPackageEntry> entries, ref long total,
        ref int nodeCount, IWorkReadOptions options, CancellationToken cancellationToken) {
        if (!entries.TryGetValue("Index.zip", out IWorkPackageEntry? nested)) return;
        using var stream = new MemoryStream(nested.Bytes, writable: false);
        ValidateZipCentralDirectory(stream, options.MaximumEntryCount - nodeCount,
            options.MaximumEntryCount, cancellationToken);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        ReadArchiveEntries(archive, entries, "Index", ref total, ref nodeCount, options,
            cancellationToken);
    }

    private static void ValidateZipCentralDirectory(Stream stream, int remainingEntries,
        int maximumEntries, CancellationToken cancellationToken) {
        OfficeArchiveSafety.ZipCentralDirectoryScanResult directory =
            OfficeArchiveSafety.ScanZipCentralDirectory(stream,
                stream.Length - stream.Position, Math.Max(0, remainingEntries), cancellationToken);
        if (!directory.IsValid) {
            throw new InvalidDataException(directory.Error
                ?? "The ZIP central directory is malformed.");
        }
        if (directory.LimitExceeded) {
            throw new InvalidDataException(
                $"Package entry count exceeds the configured limit of {maximumEntries} before package entries are opened.");
        }
    }

    private static void ReadArchiveEntries(ZipArchive archive, Dictionary<string, IWorkPackageEntry> entries,
        string? prefix, ref long total, ref int nodeCount, IWorkReadOptions options,
        CancellationToken cancellationToken) {
        foreach (ZipArchiveEntry entry in archive.Entries) {
            cancellationToken.ThrowIfCancellationRequested();
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
            byte[] bytes = ReadBounded(input, Math.Min(options.MaximumEntryBytes, entry.Length),
                normalized, cancellationToken);
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

    private static byte[] ReadBounded(Stream stream, long maximumBytes, string label,
        CancellationToken cancellationToken) {
        using var output = new MemoryStream();
        var buffer = new byte[81920];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
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
