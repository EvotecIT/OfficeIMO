using OfficeIMO.Email;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeIMO.Email.Store;

internal sealed class MailboxDirectoryStoreSessionBackend : IEmailStoreSessionBackend {
    private readonly string _root;
    private readonly string _unixOpenRoot;
    private readonly string _windowsOpenRoot;
    private readonly StringComparison _rootComparison;
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
        string rootWithoutSeparator = _root.TrimEnd(Path.DirectorySeparatorChar);
        string? rootParent = Path.GetDirectoryName(rootWithoutSeparator);
        _rootComparison = EmailStorePathIdentity.GetComparison(rootParent ?? rootWithoutSeparator);
        _unixOpenRoot = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? _root
            : AppendSeparator(ResolveUnixRealPath(rootWithoutSeparator) ?? rootWithoutSeparator);
        _windowsOpenRoot = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? AppendSeparator(EmailStorePathIdentity.ResolvePhysicalPath(
                rootWithoutSeparator))
            : _root;
        _options = options;
        DisplayName = new DirectoryInfo(path).Name;
        Index(cancellationToken);
    }

    public EmailStoreFormat Format => EmailStoreFormat.MailboxDirectory;
    public string? DisplayName { get; }
    public long SourceLength => _sourceLength;
    public IReadOnlyList<EmailStoreFolderInfo> Folders => _folders;
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _diagnostics;
    internal string RootPath => _root;

    public IEnumerable<EmailStoreItemReference> EnumerateItems(
        EmailStoreEnumerationOptions options, CancellationToken cancellationToken) {
        HashSet<string>? folders = ResolveFolderIds(options);
        if (!options.IncludeRegularItems) yield break;
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
        EmailStoreItemSummary.FromItem(ReadItem(reference,
            new EmailStoreItemReadOptions(EmailStoreItemReadParts.Metadata), cancellationToken));

    public EmailStoreItem ReadItem(EmailStoreItemReference reference, EmailStoreItemReadOptions options,
        CancellationToken cancellationToken) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        if (!_filesById.TryGetValue(reference.Id, out MailboxFile? file) ||
            file.FolderId != reference.FolderId || reference.IsAssociated || reference.IsOrphaned) {
            throw new KeyNotFoundException(
                "The item reference does not belong to this mailbox-directory session.");
        }
        bool includeAttachmentContent = options.Includes(EmailStoreItemReadParts.AttachmentContent);
        using (FileStream stream = OpenRegularMailboxFile(file.Path)) {
            EmailDocument document;
            if (file.IsEmlx) {
                EmailStoreReadResult result = new EmlxStoreReader(_options, includeAttachmentContent)
                    .Read(stream, Path.GetFileName(file.Path), cancellationToken);
                foreach (EmailStoreDiagnostic diagnostic in result.Diagnostics) _diagnostics.Add(diagnostic);
                document = result.Store.Folders.SelectMany(folder => folder.Items).Single().Document;
            } else {
                EmailReadResult result = EmailStoreMessageReader.Read(stream, _options, cancellationToken,
                    includeAttachmentContent);
                CopyDiagnostics(result.Diagnostics, file.RelativePath);
                document = result.Document;
            }
            document.Properties["EmailStore:ContainerFormat"] = Format.ToString();
            document.Properties["EmailStore:ItemId"] = file.Id;
            document.Properties["EmailStore:FolderId"] = file.FolderId;
            document.Properties["EmailStore:RelativePath"] = file.RelativePath;
            ApplyMaildirFlags(document, file.MaildirFlags);
            EmailStoreItemReadParts loadedParts = EmailStoreItemReadParts.All;
            if (!includeAttachmentContent) loadedParts &= ~EmailStoreItemReadParts.AttachmentContent;
            return new EmailStoreItem(file.Id, file.FolderId, document,
                loadedParts: loadedParts, format: EmailStoreFormat.MailboxDirectory);
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

            StringComparer entryComparer = EmailStorePathIdentity.GetComparer(current.Path);
            foreach (FileSystemInfo entry in entries.OrderBy(item => item.Name, entryComparer)) {
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
                if (!IsRegularMailboxFile(file.FullName)) {
                    _diagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_DIRECTORY_SPECIAL_FILE_SKIPPED",
                        "A non-regular mailbox candidate was skipped without opening it as a blocking stream.",
                        EmailStoreDiagnosticSeverity.Warning,
                        ToRelativePath(file.FullName)));
                    continue;
                }
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
            .GroupBy(candidate => candidate.FolderPath, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        foreach (string path in folderCounts.Keys.OrderBy(item => item, StringComparer.Ordinal)) {
            EnsureFolder(path, folderCounts);
        }
        foreach (MailboxCandidate candidate in candidates.OrderBy(
            item => item.RelativePath, StringComparer.Ordinal)) {
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
        if (normalized.StartsWith(_root, _rootComparison)) {
            return normalized.Substring(_root.Length).Replace('\\', '/');
        }
        string rootWithoutSeparator = _root.TrimEnd(Path.DirectorySeparatorChar);
        return string.Equals(normalized, rootWithoutSeparator, _rootComparison)
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

    private bool IsEmlx(FileInfo file) {
        if (!file.Name.EndsWith(".emlx", StringComparison.OrdinalIgnoreCase)) return false;
        string? parent = file.Directory?.Name;
        if (!string.Equals(parent, "cur", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(parent, "new", StringComparison.OrdinalIgnoreCase)) return true;
        try {
            if (!TryOpenRegularMailboxFile(file.FullName, 4 * 1024, out FileStream stream)) {
                return false;
            }
            using (stream) {
                return EmlxStoreReader.HasEnvelopePrefix(stream);
            }
        } catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException) {
            return false;
        }
    }

    private bool IsRegularMailboxFile(string path) {
        if (!TryOpenRegularMailboxFile(path, 4 * 1024, out FileStream stream)) return false;
        stream.Dispose();
        return true;
    }

    private FileStream OpenRegularMailboxFile(string path) {
        if (TryOpenRegularMailboxFile(path, 64 * 1024, out FileStream stream)) return stream;
        throw new IOException("The mailbox item is no longer a regular readable file.");
    }

    private bool TryOpenRegularMailboxFile(
        string path,
        int bufferSize,
        out FileStream stream) {
        stream = null!;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            try {
                stream = new FileStream(
                    path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize,
                    FileOptions.SequentialScan);
                if (!OpenedPathRemainsInsideRoot(stream.SafeFileHandle, path)) {
                    stream.Dispose();
                    stream = null!;
                    return false;
                }
                return true;
            } catch (Exception exception) when (
                exception is IOException || exception is UnauthorizedAccessException) {
                return false;
            }
        }

        int nonBlocking = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? 0x0004 : 0x0800;
        int closeOnExec = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? 0x01000000 : 0x00080000;
        int noFollow = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? 0x00000100 : 0x00020000;
        int descriptor = OpenUnixPathWithoutLinks(path, nonBlocking | closeOnExec | noFollow);
        if (descriptor < 0) return false;
        if (!IsRegularUnixDescriptor(descriptor)) {
            CloseUnix(descriptor);
            return false;
        }
        if (SeekUnix(descriptor, 0L, 1) < 0L) {
            CloseUnix(descriptor);
            return false;
        }

        var handle = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
        try {
            stream = new FileStream(handle, FileAccess.Read, bufferSize, isAsync: false);
            return true;
        } catch {
            handle.Dispose();
            throw;
        }
    }

    private static bool IsRegularUnixDescriptor(int descriptor) =>
        GetUnixFileStatus(new IntPtr(descriptor), out UnixFileStatus status) == 0
        && (status.Mode & 0xF000) == 0x8000;

    private bool OpenedPathRemainsInsideRoot(SafeFileHandle handle, string requestedPath) {
        var buffer = new StringBuilder(1024);
        uint length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
        if (length == 0) return false;
        if (length >= buffer.Capacity) {
            buffer = new StringBuilder(checked((int)length + 1));
            length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
            if (length == 0 || length >= buffer.Capacity) return false;
        }
        return IsResolvedPathInsideRoot(NormalizeWindowsFinalPath(buffer.ToString()), requestedPath);
    }

    private bool IsResolvedPathInsideRoot(string resolvedPath, string requestedPath) {
        string normalized;
        try {
            normalized = Path.GetFullPath(resolvedPath);
        } catch (Exception exception) when (
            exception is ArgumentException || exception is NotSupportedException || exception is PathTooLongException) {
            return false;
        }
        string requested = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? EmailStorePathIdentity.ResolvePhysicalPath(requestedPath)
            : Path.GetFullPath(requestedPath);
        return normalized.StartsWith(_windowsOpenRoot, _rootComparison)
            && string.Equals(normalized, requested, _rootComparison);
    }

    private static string NormalizeWindowsFinalPath(string path) {
        const string uncPrefix = @"\\?\UNC\";
        const string devicePrefix = @"\\?\";
        if (path.StartsWith(uncPrefix, StringComparison.OrdinalIgnoreCase)) {
            return @"\\" + path.Substring(uncPrefix.Length);
        }
        return path.StartsWith(devicePrefix, StringComparison.OrdinalIgnoreCase)
            ? path.Substring(devicePrefix.Length)
            : path;
    }

    private int OpenUnixPathWithoutLinks(string path, int fileFlags) {
        string normalized = Path.GetFullPath(path);
        if (!normalized.StartsWith(_root, _rootComparison)) return -1;
        string relative = normalized.Substring(_root.Length);
        string[] segments = relative.Split(
            new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
            StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length == 0 || segments.Any(segment => segment == "." || segment == "..")) return -1;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            const int noFollow = 0x00000100;
            const int noFollowAny = 0x20000000;
            string canonicalPath = Path.Combine(_unixOpenRoot, string.Join(Path.DirectorySeparatorChar.ToString(), segments));
            return OpenUnix(canonicalPath, (fileFlags & ~noFollow) | noFollowAny);
        }

        const int linuxCloseOnExec = 0x00080000;
        const int linuxNoFollow = 0x00020000;
        const int linuxDirectory = 0x00010000;
        int directory = OpenUnix(
            _unixOpenRoot.TrimEnd(Path.DirectorySeparatorChar),
            linuxCloseOnExec | linuxNoFollow | linuxDirectory);
        if (directory < 0) return -1;
        try {
            for (int index = 0; index < segments.Length - 1; index++) {
                int child = OpenAtUnix(
                    directory,
                    segments[index],
                    linuxCloseOnExec | linuxNoFollow | linuxDirectory);
                if (child < 0) return -1;
                CloseUnix(directory);
                directory = child;
            }
            return OpenAtUnix(directory, segments[segments.Length - 1], fileFlags);
        } finally {
            CloseUnix(directory);
        }
    }

    private static string? ResolveUnixRealPath(string path) {
        IntPtr resolved = RealPathUnix(path, IntPtr.Zero);
        if (resolved == IntPtr.Zero) return null;
        try {
            return Marshal.PtrToStringAnsi(resolved);
        } finally {
            FreeUnix(resolved);
        }
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetFinalPathNameByHandle(
        SafeFileHandle file,
        StringBuilder filePath,
        uint filePathLength,
        uint flags);

    [DllImport("libc", EntryPoint = "open", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int OpenUnix(string path, int flags);

    [DllImport("libc", EntryPoint = "openat", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int OpenAtUnix(int directoryDescriptor, string path, int flags);

    [DllImport("libc", EntryPoint = "lseek", SetLastError = true)]
    private static extern long SeekUnix(int descriptor, long offset, int origin);

    [DllImport("System.Native", EntryPoint = "SystemNative_FStat", SetLastError = true)]
    private static extern int GetUnixFileStatus(IntPtr descriptor, out UnixFileStatus status);

    [DllImport("libc", EntryPoint = "close", SetLastError = true)]
    private static extern int CloseUnix(int descriptor);

    [StructLayout(LayoutKind.Sequential)]
    private struct UnixFileStatus {
        internal int Flags;
        internal int Mode;
        internal uint Uid;
        internal uint Gid;
        internal long Size;
        internal long AccessTime;
        internal long AccessTimeNanoseconds;
        internal long ModificationTime;
        internal long ModificationTimeNanoseconds;
        internal long ChangeTime;
        internal long ChangeTimeNanoseconds;
        internal long BirthTime;
        internal long BirthTimeNanoseconds;
        internal long Device;
        internal long RawDevice;
        internal long Inode;
        internal uint UserFlags;
        internal int HardLinkCount;
    }

    [DllImport("libc", EntryPoint = "realpath", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern IntPtr RealPathUnix(string path, IntPtr resolvedPath);

    [DllImport("libc", EntryPoint = "free")]
    private static extern void FreeUnix(IntPtr pointer);

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
