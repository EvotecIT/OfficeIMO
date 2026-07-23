using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

internal sealed class OabDiscoveryResult {
    internal OabDiscoveryResult(IReadOnlyList<OfflineAddressBookFileInfo> files,
        IReadOnlyList<OabSource> fullDetailsSources,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        Files = files;
        FullDetailsSources = fullDetailsSources;
        Diagnostics = diagnostics;
    }

    internal IReadOnlyList<OfflineAddressBookFileInfo> Files { get; }
    internal IReadOnlyList<OabSource> FullDetailsSources { get; }
    internal IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
}

internal static class OabFileDiscovery {
    internal static OabDiscoveryResult Discover(string path, OfflineAddressBookReaderOptions options,
        CancellationToken cancellationToken) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        string fullPath = Path.GetFullPath(path);
        if (File.Exists(fullPath)) return DiscoverFile(fullPath, options);
        if (!Directory.Exists(fullPath)) throw new FileNotFoundException("OAB path does not exist.", fullPath);

        var diagnostics = new List<EmailDiagnostic>();
        var paths = new List<string>();
        var pending = new Stack<DirectoryNode>();
        int inspectedEntries = 0;
        pending.Push(new DirectoryNode(fullPath, 0));
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            DirectoryNode node = pending.Pop();
            FileSystemInfo[] entries;
            try {
                int remainingEntries = options.MaxDirectoryEntries - inspectedEntries;
                int probeEntries = remainingEntries <= 0
                    ? 1
                    : remainingEntries == int.MaxValue
                        ? int.MaxValue
                        : remainingEntries + 1;
                entries = new DirectoryInfo(node.Path)
                    .EnumerateFileSystemInfos()
                    .Take(probeEntries)
                    .ToArray();
            } catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException) {
                diagnostics.Add(new EmailDiagnostic(
                    "OAB_DIRECTORY_SKIPPED",
                    exception.Message,
                    EmailDiagnosticSeverity.Warning,
                    node.Path));
                continue;
            }
            if (entries.Length > options.MaxDirectoryEntries - inspectedEntries) {
                throw new OfflineAddressBookLimitExceededException(
                    nameof(options.MaxDirectoryEntries), inspectedEntries + entries.Length,
                    options.MaxDirectoryEntries, fullPath);
            }
            inspectedEntries += entries.Length;

            foreach (FileSystemInfo entry in entries.OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase)) {
                cancellationToken.ThrowIfCancellationRequested();
                if ((entry.Attributes & FileAttributes.ReparsePoint) != 0) {
                    diagnostics.Add(new EmailDiagnostic(
                        "OAB_REPARSE_POINT_SKIPPED",
                        "A reparse point was skipped during bounded OAB discovery.",
                        EmailDiagnosticSeverity.Information,
                        entry.FullName));
                    continue;
                }
                if ((entry.Attributes & FileAttributes.Directory) != 0) {
                    if (node.Depth >= options.MaxDirectoryDepth) {
                        diagnostics.Add(new EmailDiagnostic(
                            "OAB_DIRECTORY_DEPTH_LIMIT",
                            "A directory was not traversed because the configured OAB discovery depth was reached.",
                            EmailDiagnosticSeverity.Warning,
                            entry.FullName));
                    } else {
                        pending.Push(new DirectoryNode(entry.FullName, node.Depth + 1));
                    }
                    continue;
                }
                if (!string.Equals(entry.Extension, ".oab", StringComparison.OrdinalIgnoreCase)) continue;
                if (paths.Count >= options.MaxDiscoveredFiles) {
                    throw new OfflineAddressBookLimitExceededException(
                        nameof(options.MaxDiscoveredFiles), paths.Count + 1L,
                        options.MaxDiscoveredFiles, fullPath);
                }
                paths.Add(entry.FullName);
            }
        }

        var files = new List<OfflineAddressBookFileInfo>(paths.Count);
        var sources = new List<OabSource>();
        foreach (string filePath in paths.OrderBy(value => value, StringComparer.OrdinalIgnoreCase)) {
            cancellationToken.ThrowIfCancellationRequested();
            OfflineAddressBookFileInfo info = InspectFile(filePath);
            files.Add(info);
            if (info.Format == OfflineAddressBookFormat.Version4FullDetails) {
                if (info.Length > options.MaxInputBytes) {
                    throw new OfflineAddressBookLimitExceededException(
                        nameof(options.MaxInputBytes), info.Length, options.MaxInputBytes, filePath);
                }
                sources.Add(OabSource.FromFile(filePath));
            }
        }
        return new OabDiscoveryResult(files, sources, diagnostics);
    }

    internal static OfflineAddressBookFileInfo InspectStream(OabSource source) {
        using (OabStreamLease lease = source.OpenRead()) {
            Stream stream = lease.Stream;
            uint version = source.Length >= 4
                ? OabBinary.ReadUInt32(stream, source.SourcePath)
                : 0;
            OfflineAddressBookFormat format = Classify(source.SourceName, version);
            return new OfflineAddressBookFileInfo(
                source.SourcePath, source.SourceName, source.Length, version, format);
        }
    }

    private static OabDiscoveryResult DiscoverFile(string path, OfflineAddressBookReaderOptions options) {
        OfflineAddressBookFileInfo info = InspectFile(path);
        if (info.Length > options.MaxInputBytes && info.Format == OfflineAddressBookFormat.Version4FullDetails) {
            throw new OfflineAddressBookLimitExceededException(
                nameof(options.MaxInputBytes), info.Length, options.MaxInputBytes, path);
        }
        IReadOnlyList<OabSource> sources = info.Format == OfflineAddressBookFormat.Version4FullDetails
            ? new[] { OabSource.FromFile(path) }
            : Array.Empty<OabSource>();
        return new OabDiscoveryResult(new[] { info }, sources, Array.Empty<EmailDiagnostic>());
    }

    private static OfflineAddressBookFileInfo InspectFile(string path) {
        var file = new FileInfo(path);
        uint version = 0;
        if (file.Length >= 4) {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete)) {
                version = OabBinary.ReadUInt32(stream, path);
            }
        }
        return new OfflineAddressBookFileInfo(
            file.FullName, file.Name, file.Length, version, Classify(file.Name, version));
    }

    private static OfflineAddressBookFormat Classify(string name, uint version) {
        if (version == 0x00000020U) return OfflineAddressBookFormat.Version4FullDetails;
        if (version == 0x00000007U) return OfflineAddressBookFormat.DisplayTemplate;
        string normalized = Path.GetFileName(name).ToLowerInvariant();
        if (normalized == "browse.oab" || normalized == "ubrowse.oab") return OfflineAddressBookFormat.LegacyBrowse;
        if (normalized == "anrdex.oab" || normalized == "uanrdex.oab") return OfflineAddressBookFormat.LegacyAnrIndex;
        if (normalized == "rdndex.oab" || normalized == "urdndex.oab") return OfflineAddressBookFormat.LegacyRdnIndex;
        if (normalized == "details.oab" || normalized == "udetails.oab") return OfflineAddressBookFormat.LegacyDetails;
        if (normalized == "changes.oab" || normalized == "uchanges.oab") return OfflineAddressBookFormat.LegacyChanges;
        return OfflineAddressBookFormat.Unknown;
    }

    private sealed class DirectoryNode {
        internal DirectoryNode(string path, int depth) {
            Path = path;
            Depth = depth;
        }
        internal string Path { get; }
        internal int Depth { get; }
    }
}
