using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class MailboxDirectoryStoreSessionBackend : IEmailStoreSessionBackend {
    private readonly string _root;
    private readonly StringComparer _pathComparer;
    private readonly StringComparison _pathComparison;
    private readonly EmailStoreReaderOptions _options;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private readonly List<EmailStoreFolderInfo> _folders = new List<EmailStoreFolderInfo>();
    private readonly List<MailboxFile> _files = new List<MailboxFile>();
    private readonly Dictionary<string, MailboxFile> _filesById =
        new Dictionary<string, MailboxFile>(StringComparer.Ordinal);
    private readonly Dictionary<string, EmailStoreFolderInfo> _foldersById =
        new Dictionary<string, EmailStoreFolderInfo>(StringComparer.Ordinal);
    private long _sourceLength;

    internal MailboxDirectoryStoreSessionBackend(string path, EmailStoreReaderOptions options,
        CancellationToken cancellationToken) {
        _root = AppendSeparator(Path.GetFullPath(path));
        _pathComparer = EmailStorePathIdentity.GetComparer(_root);
        _pathComparison = EmailStorePathIdentity.GetComparison(_root);
        _options = options;
        DisplayName = new DirectoryInfo(path).Name;
        Index(cancellationToken);
    }

    public EmailStoreFormat Format => EmailStoreFormat.MailboxDirectory;
    public string? DisplayName { get; }
    public long SourceLength => _sourceLength;
    public IReadOnlyList<EmailStoreFolderInfo> Folders => _folders;
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _diagnostics;

    public IEnumerable<EmailStoreItemReference> EnumerateItems(
        EmailStoreEnumerationOptions options, CancellationToken cancellationToken) {
        HashSet<string>? folders = ResolveFolderIds(options);
        int count = 0;
        foreach (MailboxFile file in _files) {
            cancellationToken.ThrowIfCancellationRequested();
            if (folders != null && !folders.Contains(file.FolderId)) continue;
            if (++count > options.MaxItems) yield break;
            yield return new EmailStoreItemReference(file.Id, file.FolderId, false, false);
        }
    }

    public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference,
        CancellationToken cancellationToken) =>
        EmailStoreItemSummary.FromItem(ReadItem(reference, EmailStoreItemReadOptions.Default, cancellationToken));

    public EmailStoreItem ReadItem(EmailStoreItemReference reference, EmailStoreItemReadOptions options,
        CancellationToken cancellationToken) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        if (!_filesById.TryGetValue(reference.Id, out MailboxFile? file) ||
            file.FolderId != reference.FolderId || reference.IsAssociated || reference.IsOrphaned) {
            throw new KeyNotFoundException(
                "The item reference does not belong to this mailbox-directory session.");
        }
        using (var stream = new FileStream(
            file.Path, FileMode.Open, FileAccess.Read, FileShare.Read, 64 * 1024, FileOptions.SequentialScan)) {
            EmailDocument document;
            if (file.IsEmlx) {
                EmailStoreReadResult result = new EmlxStoreReader(_options)
                    .Read(stream, Path.GetFileName(file.Path), cancellationToken);
                foreach (EmailStoreDiagnostic diagnostic in result.Diagnostics) _diagnostics.Add(diagnostic);
                document = result.Store.Folders.SelectMany(folder => folder.Items).Single().Document;
            } else {
                EmailReadResult result = EmailStoreMessageReader.Read(stream, _options, cancellationToken);
                CopyDiagnostics(result.Diagnostics, file.RelativePath);
                document = result.Document;
            }
            document.Properties["EmailStore:ContainerFormat"] = Format.ToString();
            document.Properties["EmailStore:ItemId"] = file.Id;
            document.Properties["EmailStore:FolderId"] = file.FolderId;
            document.Properties["EmailStore:RelativePath"] = file.RelativePath;
            ApplyMaildirFlags(document, file.MaildirFlags);
            return new EmailStoreItem(
                file.Id, file.FolderId, document, format: EmailStoreFormat.MailboxDirectory);
        }
    }

    public void Dispose() { }

    private void Index(CancellationToken cancellationToken) {
        var candidates = new List<MailboxCandidate>();
        var pending = new Stack<DirectoryCandidate>();
        pending.Push(new DirectoryCandidate(_root.TrimEnd(Path.DirectorySeparatorChar), 0));
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            DirectoryCandidate current = pending.Pop();
            FileSystemInfo[] entries;
            try {
                entries = new DirectoryInfo(current.Path).GetFileSystemInfos();
            } catch (Exception exception) when (
                exception is IOException || exception is UnauthorizedAccessException) {
                _diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_DIRECTORY_ENUMERATION_FAILED",
                    exception.Message,
                    EmailStoreDiagnosticSeverity.Warning,
                    ToRelativePath(current.Path)));
                continue;
            }

            foreach (FileSystemInfo entry in entries.OrderBy(item => item.Name, _pathComparer)) {
                cancellationToken.ThrowIfCancellationRequested();
                if ((entry.Attributes & FileAttributes.ReparsePoint) != 0) {
                    _diagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_DIRECTORY_REPARSE_POINT_SKIPPED",
                        "A symbolic link or reparse point was skipped to keep traversal inside the mailbox root.",
                        EmailStoreDiagnosticSeverity.Information,
                        ToRelativePath(entry.FullName)));
                    continue;
                }
                if (entry is DirectoryInfo directory) {
                    if (current.Depth >= _options.MaxDirectoryDepth) {
                        throw new EmailStoreLimitExceededException(
                            nameof(EmailStoreReaderOptions.MaxDirectoryDepth),
                            current.Depth + 1L,
                            _options.MaxDirectoryDepth);
                    }
                    pending.Push(new DirectoryCandidate(directory.FullName, current.Depth + 1));
                    continue;
                }
                if (!(entry is FileInfo file) || !IsMailboxFile(file)) continue;
                if (candidates.Count >= _options.MaxDirectoryFileCount) {
                    throw new EmailStoreLimitExceededException(
                        nameof(EmailStoreReaderOptions.MaxDirectoryFileCount),
                        candidates.Count + 1L,
                        _options.MaxDirectoryFileCount);
                }
                _sourceLength = AddBounded(_sourceLength, file.Length);
                candidates.Add(new MailboxCandidate(
                    file.FullName,
                    ToRelativePath(file.FullName),
                    IsEmlx(file),
                    GetLogicalFolderPath(ToRelativePath(file.DirectoryName ?? _root)),
                    ParseMaildirFlags(file.Name, file.Directory?.Name)));
            }
        }

        var folderCounts = candidates
            .GroupBy(candidate => candidate.FolderPath, _pathComparer)
            .ToDictionary(group => group.Key, group => group.Count(), _pathComparer);
        foreach (string path in folderCounts.Keys.OrderBy(item => item, _pathComparer)) {
            EnsureFolder(path, folderCounts);
        }
        foreach (MailboxCandidate candidate in candidates.OrderBy(
            item => item.RelativePath, _pathComparer)) {
            string folderId = GetFolderId(candidate.FolderPath);
            string id = string.Concat("directory:item:", candidate.RelativePath.Replace('\\', '/'));
            var file = new MailboxFile(
                id, folderId, candidate.Path, candidate.RelativePath, candidate.IsEmlx,
                candidate.MaildirFlags);
            _files.Add(file);
            _filesById.Add(id, file);
        }
    }

    private void EnsureFolder(string path, IReadOnlyDictionary<string, int> counts) {
        if (_foldersById.ContainsKey(GetFolderId(path))) return;
        string? parentPath = GetParentPath(path);
        if (parentPath != null) EnsureFolder(parentPath, counts);
        string id = GetFolderId(path);
        string? parentId = parentPath == null ? null : GetFolderId(parentPath);
        string name = path == "." ? (DisplayName ?? "Mailbox") : GetLastPart(path);
        int count = counts.TryGetValue(path, out int directCount) ? directCount : 0;
        var folder = new EmailStoreFolderInfo(id, parentId, name, count, 0);
        _folders.Add(folder);
        _foldersById.Add(id, folder);
    }

    private HashSet<string>? ResolveFolderIds(EmailStoreEnumerationOptions options) {
        if (options.FolderId == null) return null;
        if (!_foldersById.ContainsKey(options.FolderId)) {
            throw new KeyNotFoundException(
                "The requested folder does not belong to this mailbox-directory session.");
        }
        var result = new HashSet<string>(StringComparer.Ordinal) { options.FolderId };
        if (!options.IncludeDescendants) return result;
        bool added;
        do {
            added = false;
            foreach (EmailStoreFolderInfo folder in _folders) {
                if (folder.ParentId != null && result.Contains(folder.ParentId) && result.Add(folder.Id)) {
                    added = true;
                }
            }
        } while (added);
        return result;
    }

    private bool IsMailboxFile(FileInfo file) {
        string extension = file.Extension;
        if (extension.Equals(".emlx", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".eml", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".mime", StringComparison.OrdinalIgnoreCase)) return true;
        string? parent = file.Directory?.Name;
        return string.Equals(parent, "cur", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(parent, "new", StringComparison.OrdinalIgnoreCase);
    }

    private long AddBounded(long current, long length) {
        if (length < 0 || current > _options.MaxInputBytes - length) {
            long actual = length > long.MaxValue - current ? long.MaxValue : current + length;
            throw new EmailStoreLimitExceededException(
                nameof(EmailStoreReaderOptions.MaxInputBytes), actual, _options.MaxInputBytes);
        }
        return current + length;
    }

    private string ToRelativePath(string fullPath) {
        string normalized = Path.GetFullPath(fullPath);
        if (normalized.StartsWith(_root, _pathComparison)) {
            return normalized.Substring(_root.Length).Replace('\\', '/');
        }
        string rootWithoutSeparator = _root.TrimEnd(Path.DirectorySeparatorChar);
        return string.Equals(normalized, rootWithoutSeparator, _pathComparison)
            ? string.Empty
            : normalized;
    }

    private static string GetLogicalFolderPath(string relativeDirectory) {
        string[] parts = relativeDirectory.Replace('\\', '/')
            .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        string[] mailboxParts = parts
            .Where(part => part.EndsWith(".mbox", StringComparison.OrdinalIgnoreCase))
            .Select(part => part.Substring(0, part.Length - 5))
            .Where(part => part.Length > 0)
            .ToArray();
        if (mailboxParts.Length > 0) return string.Join("/", mailboxParts);
        int length = parts.Length;
        if (length > 0 && (string.Equals(parts[length - 1], "cur", StringComparison.OrdinalIgnoreCase) ||
                           string.Equals(parts[length - 1], "new", StringComparison.OrdinalIgnoreCase) ||
                           string.Equals(parts[length - 1], "Messages", StringComparison.OrdinalIgnoreCase))) {
            length--;
        }
        if (length == 0) return ".";
        string[] visible = parts.Take(length)
            .Select(part => part.Length > 1 && part[0] == '.' ? part.Substring(1) : part)
            .ToArray();
        return string.Join("/", visible);
    }

    private void CopyDiagnostics(IEnumerable<EmailDiagnostic> diagnostics, string location) {
        foreach (EmailDiagnostic diagnostic in diagnostics) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.Severity == EmailDiagnosticSeverity.Error
                    ? EmailStoreDiagnosticSeverity.Error
                    : diagnostic.Severity == EmailDiagnosticSeverity.Information
                        ? EmailStoreDiagnosticSeverity.Information
                        : EmailStoreDiagnosticSeverity.Warning,
                diagnostic.Location == null ? location : string.Concat(location, "/", diagnostic.Location)));
        }
    }

    private static bool IsEmlx(FileInfo file) {
        string? parent = file.Directory?.Name;
        if (string.Equals(parent, "cur", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(parent, "new", StringComparison.OrdinalIgnoreCase)) return false;
        return file.Name.EndsWith(".emlx", StringComparison.OrdinalIgnoreCase);
    }

    internal static string? ParseMaildirFlags(string name, string? parentDirectoryName) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        if (!string.Equals(parentDirectoryName, "cur", StringComparison.OrdinalIgnoreCase)) return null;
        int marker = name.LastIndexOf(":2,", StringComparison.Ordinal);
        if (marker <= 0) return null;
        string flags = name.Substring(marker + 3);
        for (int index = 0; index < flags.Length; index++) {
            char value = flags[index];
            if (value < 'A' || value > 'Z') return null;
        }
        return flags;
    }

    internal static void ApplyMaildirFlags(EmailDocument document, string? flags) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (flags == null) return;
        document.MessageMetadata.IsDraft = flags.IndexOf('D') >= 0;
        document.MessageMetadata.IsRead = flags.IndexOf('S') >= 0;
        document.Properties["Emlx:Flag:Flagged"] = flags.IndexOf('F') >= 0;
        document.Properties["Emlx:Flag:Forwarded"] = flags.IndexOf('P') >= 0;
        document.Properties["Emlx:Flag:Answered"] = flags.IndexOf('R') >= 0;
        document.Properties["Emlx:Flag:Deleted"] = flags.IndexOf('T') >= 0;
    }

    private static string GetFolderId(string path) =>
        string.Concat("directory:folder:", path);

    private static string? GetParentPath(string path) {
        if (path == ".") return null;
        int slash = path.LastIndexOf('/');
        return slash < 0 ? null : path.Substring(0, slash);
    }

    private static string GetLastPart(string path) {
        int slash = path.LastIndexOf('/');
        return slash < 0 ? path : path.Substring(slash + 1);
    }

    private static string AppendSeparator(string path) =>
        path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            ? path
            : string.Concat(path, Path.DirectorySeparatorChar.ToString());

    private sealed class DirectoryCandidate {
        internal DirectoryCandidate(string path, int depth) { Path = path; Depth = depth; }
        internal string Path { get; }
        internal int Depth { get; }
    }

    private sealed class MailboxCandidate {
        internal MailboxCandidate(string path, string relativePath, bool isEmlx, string folderPath,
            string? maildirFlags) {
            Path = path;
            RelativePath = relativePath;
            IsEmlx = isEmlx;
            FolderPath = folderPath;
            MaildirFlags = maildirFlags;
        }
        internal string Path { get; }
        internal string RelativePath { get; }
        internal bool IsEmlx { get; }
        internal string FolderPath { get; }
        internal string? MaildirFlags { get; }
    }

    private sealed class MailboxFile {
        internal MailboxFile(string id, string folderId, string path, string relativePath, bool isEmlx,
            string? maildirFlags) {
            Id = id;
            FolderId = folderId;
            Path = path;
            RelativePath = relativePath;
            IsEmlx = isEmlx;
            MaildirFlags = maildirFlags;
        }
        internal string Id { get; }
        internal string FolderId { get; }
        internal string Path { get; }
        internal string RelativePath { get; }
        internal bool IsEmlx { get; }
        internal string? MaildirFlags { get; }
    }
}
