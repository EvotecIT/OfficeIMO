using OfficeIMO.Email;
using System.Runtime.InteropServices;

namespace OfficeIMO.Email.Store;

internal sealed partial class MailboxDirectoryStoreSessionBackend : IEmailStoreSessionBackend {
    private readonly string _root;
    private readonly string _unixOpenRoot;
    private readonly string _windowsOpenRoot;
    private readonly StringComparison _rootComparison;
    private readonly EmailStoreReaderOptions _options;
    private readonly EmailStoreReadResources _resources;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private readonly List<EmailStoreFolderInfo> _folders = new List<EmailStoreFolderInfo>();
    private readonly List<MailboxFile> _files = new List<MailboxFile>();
    private readonly List<EmailStoreItemReference> _items = new List<EmailStoreItemReference>();
    private readonly HashSet<string> _diagnosticKeys = new HashSet<string>(StringComparer.Ordinal);
    private readonly Dictionary<string, AggregateItem> _aggregateItems =
        new Dictionary<string, AggregateItem>(StringComparer.Ordinal);
    private readonly Dictionary<string, MailboxFile> _filesById =
        new Dictionary<string, MailboxFile>(StringComparer.Ordinal);
    private readonly Dictionary<string, List<MailboxFile>> _partialFiles;
    private readonly Dictionary<string, EmailStoreFolderInfo> _foldersById =
        new Dictionary<string, EmailStoreFolderInfo>(StringComparer.Ordinal);
    private long _sourceLength;

    internal MailboxDirectoryStoreSessionBackend(string path, EmailStoreReaderOptions options,
        CancellationToken cancellationToken) {
        _root = AppendSeparator(Path.GetFullPath(path));
        string rootWithoutSeparator = TrimTrailingDirectorySeparators(_root);
        string? rootParent = Path.GetDirectoryName(rootWithoutSeparator);
        _rootComparison = EmailStorePathIdentity.GetComparison(rootParent ?? rootWithoutSeparator);
        _partialFiles = new Dictionary<string, List<MailboxFile>>(EmailStorePathIdentity.GetComparer(rootParent ?? rootWithoutSeparator));
        _unixOpenRoot = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? _root
            : AppendSeparator(ResolveUnixRealPath(rootWithoutSeparator) ?? rootWithoutSeparator);
        _windowsOpenRoot = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? AppendSeparator(EmailStorePathIdentity.ResolvePhysicalPath(
                rootWithoutSeparator))
            : _root;
        _options = options;
        _resources = new EmailStoreReadResources(options);
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
        foreach (EmailStoreItemReference item in _items) {
            cancellationToken.ThrowIfCancellationRequested();
            if (folders != null && !folders.Contains(item.FolderId)) continue;
            if (++count > options.MaxItems) yield break;
            yield return item;
        }
    }

    public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (TryGetAggregate(reference, out AggregateItem aggregate)) {
            return aggregate.Backend.ReadSummary(aggregate.Reference, cancellationToken);
        }
        if (!_filesById.TryGetValue(reference.Id, out MailboxFile? file) || file.FolderId != reference.FolderId ||
            reference.IsAssociated || reference.IsOrphaned) throw new KeyNotFoundException("The item reference does not belong to this mailbox-directory session.");
        if (file.Summary != null) {
            using FileStream input = OpenRegularMailboxFile(file.Path);
            ValidateFileSource(file, input, cancellationToken);
            return file.Summary;
        }
        file.Summary = EmailStoreItemSummary.FromItem(ReadItem(reference,
            new EmailStoreItemReadOptions(EmailStoreItemReadParts.Metadata), cancellationToken));
        return file.Summary;
    }

    public EmailStoreItem ReadItem(EmailStoreItemReference reference, EmailStoreItemReadOptions options,
        CancellationToken cancellationToken) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        if (TryGetAggregate(reference, out AggregateItem aggregate)) {
            EmailStoreItem item = aggregate.Backend.ReadItem(aggregate.Reference, options, cancellationToken);
            ApplyDirectoryProperties(item.Document, reference.Id, reference.FolderId, aggregate.RelativePath);
            foreach (EmailStoreDiagnostic diagnostic in aggregate.Backend.Diagnostics) {
                AddDiagnostic(diagnostic, aggregate.RelativePath);
            }
            return new EmailStoreItem(reference.Id, reference.FolderId, item.Document,
                loadedParts: item.LoadedParts, format: Format, summary: aggregate.Backend.ReadSummary(aggregate.Reference, cancellationToken));
        }
        if (!_filesById.TryGetValue(reference.Id, out MailboxFile? file) ||
            file.FolderId != reference.FolderId || reference.IsAssociated || reference.IsOrphaned) {
            throw new KeyNotFoundException(
                "The item reference does not belong to this mailbox-directory session.");
        }
        bool includeAttachmentContent = options.Includes(EmailStoreItemReadParts.AttachmentContent);
        bool includeEmbeddedMessages = options.Includes(EmailStoreItemReadParts.EmbeddedItems);
        using (FileStream stream = OpenRegularMailboxFile(file.Path)) {
            ValidateFileSource(file, stream, cancellationToken);
            EmailDocument document;
            if (file.IsEmlx) {
                EmailStoreReadResult result = new EmlxStoreReader(_options, includeAttachmentContent, options.MaxDecodedPropertyBytes,
                    includeEmbeddedMessages, _resources, options.PreferStreamingAttachmentContent,
                    part => OpenPartialPart(file, part, cancellationToken))
                    .Read(stream, Path.GetFileName(file.Path), cancellationToken);
                ValidateFileSource(file, stream, cancellationToken);
                foreach (EmailStoreDiagnostic diagnostic in result.Diagnostics) AddDiagnostic(diagnostic, file.RelativePath);
                document = result.Store.Folders.SelectMany(folder => folder.Items).Single().Document;
            } else {
                using EmailReadResult result = EmailStoreMessageReader.Read(stream, _options, cancellationToken,
                    includeAttachmentContent, options.MaxDecodedPropertyBytes, includeEmbeddedMessages, options.PreferStreamingAttachmentContent);
                CopyDiagnostics(result.Diagnostics, file.RelativePath);
                document = result.Document;
                ValidateFileSource(file, stream, cancellationToken);
                _resources.Adopt(result);
            }
            ApplyDirectoryProperties(document, file.Id, file.FolderId, file.RelativePath);
            ApplyMaildirFlags(document, file.MaildirFlags);
            EmailStoreItemReadParts loadedParts = EmailStoreItemReadParts.All;
            if (!includeAttachmentContent) loadedParts &= ~EmailStoreItemReadParts.AttachmentContent;
            if (!includeEmbeddedMessages) loadedParts &= ~EmailStoreItemReadParts.EmbeddedItems;
            return new EmailStoreItem(file.Id, file.FolderId, document,
                loadedParts: loadedParts, format: EmailStoreFormat.MailboxDirectory);
        }
    }

    public void Dispose() => _resources.Dispose();

    private void ValidateFileSource(MailboxFile file, Stream input, CancellationToken cancellationToken) {
        // Directory catalogs remain lazy: pin each selected file at its first projection.
        if (file.SourceGuard == null) file.SourceGuard = new EmailStoreSourceGuard(input, file.Length,
            _options.MaxInputBytes, _resources.Dispose, cancellationToken);
        else file.SourceGuard.Validate(input, cancellationToken);
    }

    private Stream? OpenPartialPart(MailboxFile message, string part, CancellationToken cancellationToken) {
        string name = Path.GetFileName(message.Path);
        const string suffix = ".partial.emlx";
        if (!name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) return null;
        string identifier = name.Substring(0, name.Length - suffix.Length);
        if (identifier.Length == 0 || identifier.Any(value => value < '0' || value > '9')) return null;
        string directory = Path.GetDirectoryName(message.Path)!;
        if (Path.GetFileName(directory) != "Messages") return null;
        string path = Path.Combine(Path.GetDirectoryName(directory)!, "Attachments", identifier, part);
        // Only files captured inside the caller-selected root can supply a part; a MIME filename never becomes a path.
        if (!_partialFiles.TryGetValue(path, out List<MailboxFile>? files) || files.Count != 1) return null;
        FileStream input = OpenRegularMailboxFile(files[0].Path);
        try {
            ValidateFileSource(files[0], input, cancellationToken);
            return new EmailStoreValidatedReadStream(input, files[0].SourceGuard!, cancellationToken);
        }
        catch { input.Dispose(); throw; }
    }

    private bool TryGetAggregate(EmailStoreItemReference reference, out AggregateItem aggregate) {
        aggregate = null!;
        if (!_aggregateItems.TryGetValue(reference.Id, out AggregateItem? found)) return false;
        aggregate = found;
        if (reference.FolderId != aggregate.FolderId || reference.IsAssociated || reference.IsOrphaned) {
            throw new KeyNotFoundException("The item reference does not belong to this mailbox-directory session.");
        }
        return true;
    }

    private void ApplyDirectoryProperties(EmailDocument document, string id, string folderId, string relativePath) {
        document.Properties["EmailStore:ContainerFormat"] = Format.ToString();
        document.Properties["EmailStore:ItemId"] = id;
        document.Properties["EmailStore:FolderId"] = folderId;
        document.Properties["EmailStore:RelativePath"] = relativePath;
    }

    private void AddDiagnostic(EmailStoreDiagnostic diagnostic, string? source = null) {
        const int maximum = 10_000;
        if (_diagnostics.Count > maximum) return;
        string? location = source == null ? diagnostic.Location : diagnostic.Location == null
            ? source : string.Concat(source, "/", diagnostic.Location);
        string key = string.Concat(diagnostic.Code.Length, ":", diagnostic.Code,
            diagnostic.Message.Length, ":", diagnostic.Message, ":", (int)diagnostic.Severity, ":", location);
        if (!_diagnosticKeys.Add(key)) return;
        if (_diagnostics.Count == maximum) {
            _diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_DIAGNOSTICS_TRUNCATED",
                "Additional mailbox-directory diagnostics were omitted after the 10,000-entry limit.",
                EmailStoreDiagnosticSeverity.Warning));
            _diagnosticKeys.Clear();
            return;
        }
        _diagnostics.Add(new EmailStoreDiagnostic(diagnostic.Code, diagnostic.Message, diagnostic.Severity,
            location, diagnostic.Operation, diagnostic.ByteOffset, diagnostic.LimitName,
            diagnostic.ActualValue, diagnostic.MaximumValue, diagnostic.Disposition,
            diagnostic.DataLossRisk, diagnostic.SuggestedAction, diagnostic.IsRetryable));
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
        string rootWithoutSeparator = TrimTrailingDirectorySeparators(_root);
        return string.Equals(normalized, rootWithoutSeparator, _rootComparison)
            ? string.Empty
            : normalized;
    }

    private static string GetLogicalFolderPath(string relativeDirectory) {
        string[] parts = relativeDirectory.Replace('\\', '/')
            .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        int firstMailbox = Array.FindIndex(parts, part => part.EndsWith(".mbox", StringComparison.OrdinalIgnoreCase));
        string[] mailboxParts = parts.Skip(firstMailbox < 0 ? parts.Length : firstMailbox)
            .Where(part => part.EndsWith(".mbox", StringComparison.OrdinalIgnoreCase))
            .Select(part => part.Substring(0, part.Length - 5))
            .Where(part => part.Length > 0)
            .ToArray();
        if (mailboxParts.Length > 0) return string.Join("/", parts.Take(firstMailbox).Concat(mailboxParts));
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
            AddDiagnostic(new EmailStoreDiagnostic(
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
        internal DirectoryCandidate(string path, int depth, bool isAttachmentStorage = false) {
            Path = path; Depth = depth; IsAttachmentStorage = isAttachmentStorage;
        }
        internal string Path { get; }
        internal int Depth { get; }
        internal bool IsAttachmentStorage { get; }
    }

    private sealed class MailboxCandidate {
        internal MailboxCandidate(string path, string relativePath, bool isEmlx, string folderPath,
            string? maildirFlags, bool isAggregate = false, bool isAttachmentStorage = false) {
            Path = path;
            RelativePath = relativePath;
            IsEmlx = isEmlx;
            FolderPath = folderPath;
            MaildirFlags = maildirFlags;
            IsAggregate = isAggregate;
            IsAttachmentStorage = isAttachmentStorage;
        }
        internal string Path { get; }
        internal string RelativePath { get; }
        internal bool IsEmlx { get; }
        internal string FolderPath { get; }
        internal string? MaildirFlags { get; }
        internal bool IsAggregate { get; }
        internal bool IsAttachmentStorage { get; }
    }

    private sealed class AggregateItem {
        internal AggregateItem(MboxStoreSessionBackend backend, EmailStoreItemReference reference,
            string folderId, string relativePath) {
            Backend = backend;
            Reference = reference;
            FolderId = folderId;
            RelativePath = relativePath;
        }
        internal MboxStoreSessionBackend Backend { get; }
        internal EmailStoreItemReference Reference { get; }
        internal string FolderId { get; }
        internal string RelativePath { get; }
    }

    private sealed class MailboxFile {
        internal MailboxFile(string id, string folderId, string path, string relativePath, bool isEmlx,
            string? maildirFlags) {
            Id = id;
            FolderId = folderId;
            Path = path;
            Length = new FileInfo(path).Length;
            RelativePath = relativePath;
            IsEmlx = isEmlx;
            MaildirFlags = maildirFlags;
        }
        internal string Id { get; }
        internal string FolderId { get; }
        internal string Path { get; }
        internal long Length { get; }
        internal EmailStoreItemSummary? Summary { get; set; }
        internal EmailStoreSourceGuard? SourceGuard { get; set; }
        internal string RelativePath { get; }
        internal bool IsEmlx { get; }
        internal string? MaildirFlags { get; }
    }
}
