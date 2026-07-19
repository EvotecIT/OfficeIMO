namespace OfficeIMO.Email.Store;

/// <summary>
/// Keeps an email-store source open for bounded folder discovery, lightweight item enumeration,
/// and explicit item reads. Sessions are not thread-safe.
/// </summary>
public sealed partial class EmailStoreSession : IDisposable {
    private readonly Stream _stream;
    private readonly bool _leaveOpen;
    private readonly long _originalPosition;
    private readonly EmailStoreReaderOptions _options;
    private readonly IEmailStoreSessionBackend _backend;
    private readonly EmailStoreFolderCatalog _folderCatalog;
    private bool _disposed;

    private EmailStoreSession(Stream stream, bool leaveOpen, long originalPosition,
        EmailStoreReaderOptions options, IEmailStoreSessionBackend backend) {
        _stream = stream;
        _leaveOpen = leaveOpen;
        _originalPosition = originalPosition;
        _options = options;
        _backend = backend;
        _folderCatalog = new EmailStoreFolderCatalog(backend.Folders);
    }

    /// <summary>Detected store format.</summary>
    public EmailStoreFormat Format { get { ThrowIfDisposed(); return _backend.Format; } }

    /// <summary>Display name declared by the source when available.</summary>
    public string? DisplayName { get { ThrowIfDisposed(); return _backend.DisplayName; } }

    /// <summary>Validated source length.</summary>
    public long SourceLength { get { ThrowIfDisposed(); return _backend.SourceLength; } }

    /// <summary>Lightweight folder catalog. PST/OST item payloads are not read to build it.</summary>
    public IReadOnlyList<EmailStoreFolderInfo> Folders {
        get { ThrowIfDisposed(); return _backend.Folders; }
    }

    /// <summary>Indexed folder navigation for this session.</summary>
    public EmailStoreFolderCatalog FolderCatalog {
        get { ThrowIfDisposed(); return _folderCatalog; }
    }

    /// <summary>Structured diagnostics emitted while opening or reading the session.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics {
        get { ThrowIfDisposed(); return _backend.Diagnostics; }
    }

    internal bool IsPstPasswordProtected {
        get {
            ThrowIfDisposed();
            return _backend is PstStoreSessionBackend pst && pst.IsPasswordProtected;
        }
    }

    internal bool IsOfficeImoWriterStore {
        get {
            ThrowIfDisposed();
            return _backend is PstStoreSessionBackend pst && pst.IsOfficeImoWriterStore;
        }
    }

    /// <summary>Opens a file with random-access sharing suitable for large PST/OST sources.</summary>
    public static EmailStoreSession Open(string path, EmailStoreReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (Directory.Exists(path)) {
            EmailStoreReaderOptions effective = options ?? EmailStoreReaderOptions.Default;
            var backend = new MailboxDirectoryStoreSessionBackend(path, effective, cancellationToken);
            return new EmailStoreSession(
                Stream.Null, leaveOpen: true, originalPosition: 0, effective, backend);
        }
        var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            64 * 1024, FileOptions.RandomAccess);
        try {
            return OpenCore(stream, Path.GetFileName(path), options ?? EmailStoreReaderOptions.Default,
                leaveOpen: false, originalPosition: 0, cancellationToken);
        } catch {
            stream.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Opens a readable, seekable caller stream. Its original position is restored when the session is disposed.
    /// </summary>
    public static EmailStoreSession Open(Stream stream, string? sourceName = null,
        EmailStoreReaderOptions? options = null, bool leaveOpen = true,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead || !stream.CanSeek) {
            throw new ArgumentException("Email-store streams must be readable and seekable.", nameof(stream));
        }
        long originalPosition = stream.Position;
        return OpenCore(stream, sourceName, options ?? EmailStoreReaderOptions.Default,
            leaveOpen, originalPosition, cancellationToken);
    }

    /// <summary>Streams lightweight item references according to the requested folder and recovery scope.</summary>
    public IEnumerable<EmailStoreItemReference> EnumerateItems(
        EmailStoreEnumerationOptions? options = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        foreach (EmailStoreItemReference reference in _backend.EnumerateItems(
            options ?? new EmailStoreEnumerationOptions(), cancellationToken)) {
            ThrowIfDisposed();
            yield return reference;
        }
    }

    /// <summary>Streams lightweight references from one typed folder scope.</summary>
    public IEnumerable<EmailStoreItemReference> EnumerateItems(EmailStoreFolderId folderId,
        bool includeDescendants = false, bool includeAssociatedItems = false,
        int maxItems = int.MaxValue, CancellationToken cancellationToken = default) =>
        EnumerateItems(EmailStoreEnumerationOptions.ForFolder(folderId, includeDescendants,
            includeAssociatedItems, maxItems), cancellationToken);

    /// <summary>Reads and projects one explicitly selected item.</summary>
    public EmailStoreItem ReadItem(EmailStoreItemReference reference,
        CancellationToken cancellationToken = default) {
        return ReadItem(reference, EmailStoreItemReadOptions.Default, cancellationToken);
    }

    /// <summary>Reads only the requested parts of one explicitly selected item when the backend supports it.</summary>
    public EmailStoreItem ReadItem(EmailStoreItemReference reference,
        EmailStoreItemReadOptions options,
        CancellationToken cancellationToken = default) {
        if (reference == null) throw new ArgumentNullException(nameof(reference));
        if (options == null) throw new ArgumentNullException(nameof(options));
        ThrowIfDisposed();
        return _backend.ReadItem(reference, options, cancellationToken);
    }

    /// <summary>
    /// Reads only the small set of source properties needed for browsing and search when the format supports it.
    /// A summary already carried by <paramref name="reference"/> is returned without another source read.
    /// </summary>
    public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference,
        CancellationToken cancellationToken = default) {
        if (reference == null) throw new ArgumentNullException(nameof(reference));
        ThrowIfDisposed();
        return reference.Summary ?? _backend.ReadSummary(reference, cancellationToken);
    }

    /// <summary>
    /// Searches bounded lightweight summaries without materializing message bodies, recipients, or attachments.
    /// </summary>
    public IEnumerable<EmailStoreSearchResult> Search(EmailStoreQuery query,
        CancellationToken cancellationToken = default) {
        if (query == null) throw new ArgumentNullException(nameof(query));
        ThrowIfDisposed();
        var enumeration = new EmailStoreEnumerationOptions(
            query.FolderId,
            query.IncludeDescendants,
            query.IncludeAssociatedItems,
            query.IncludeOrphanedItems,
            query.MaxItemsScanned);
        int resultCount = 0;
        foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            EmailStoreItemSummary summary = ReadSummary(reference, cancellationToken);
            if (!Matches(query, summary)) continue;
            yield return new EmailStoreSearchResult(reference, summary);
            if (++resultCount >= query.MaxResults) yield break;
        }
    }

    /// <summary>
    /// Materializes the configured store scope as a compatibility convenience. Attachment payloads are omitted
    /// when <see cref="EmailStoreReaderOptions.RetainAttachmentContent"/> is false; use selective session reads for
    /// deferred attachment streams.
    /// </summary>
    public EmailStoreReadResult ReadAll(CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        var store = new EmailStore { Format = Format, DisplayName = DisplayName };
        var folders = new Dictionary<string, EmailStoreFolder>(StringComparer.Ordinal);
        foreach (EmailStoreFolderInfo info in Folders) {
            var folder = new EmailStoreFolder(info.Id, info.ParentId, info.Name,
                info.SpecialFolderKind, info.ClassificationSource,
                info.ContainerClass, info.IsSearchFolder, info.MapiProperties);
            folders.Add(info.Id, folder);
            store.MutableFolders.Add(folder);
        }

        var enumerationOptions = new EmailStoreEnumerationOptions(
            includeAssociatedItems: _options.IncludeAssociatedItems,
            includeOrphanedItems: _options.IncludeOrphanedItems,
            maxItems: _options.MaxItemCount == int.MaxValue ? int.MaxValue : _options.MaxItemCount + 1);
        EmailStoreItemReadOptions materializationOptions = _options.RetainAttachmentContent
            ? EmailStoreItemReadOptions.Default
            : new EmailStoreItemReadOptions(
                EmailStoreItemReadParts.All & ~EmailStoreItemReadParts.AttachmentContent);
        int itemCount = 0;
        long totalAttachmentBytes = 0;
        foreach (EmailStoreItemReference reference in EnumerateItems(
            enumerationOptions, cancellationToken)) {
            itemCount++;
            if (itemCount > _options.MaxItemCount) {
                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxItemCount),
                    itemCount, _options.MaxItemCount);
            }
            EmailStoreItem item = ReadItem(reference, materializationOptions, cancellationToken);
            if (_options.RetainAttachmentContent) {
                totalAttachmentBytes = EmailStoreAttachmentBudget.AddDocument(
                    item.Document, totalAttachmentBytes, _options.MaxTotalAttachmentBytes);
            }
            EmailStoreFolder folder = folders[reference.FolderId];
            if (reference.IsAssociated) folder.MutableAssociatedItems.Add(item);
            else folder.MutableItems.Add(item);
        }
        return new EmailStoreReadResult(store, Diagnostics.ToArray(), SourceLength);
    }

    /// <summary>Closes owned sources or restores the position of a caller-owned stream.</summary>
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _backend.Dispose();
        if (_leaveOpen) {
            if (_stream.CanSeek) _stream.Position = _originalPosition;
        } else {
            _stream.Dispose();
        }
    }

    private static EmailStoreSession OpenCore(Stream stream, string? sourceName,
        EmailStoreReaderOptions options, bool leaveOpen, long originalPosition,
        CancellationToken cancellationToken) {
        if (stream.Length > options.MaxInputBytes) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxInputBytes),
                stream.Length, options.MaxInputBytes);
        }
        EmailStoreFormat format = EmailStoreReader.DetectFormat(stream, sourceName);
        stream.Position = 0;
        try {
            IEmailStoreSessionBackend backend;
            switch (format) {
                case EmailStoreFormat.Pst:
                case EmailStoreFormat.Ost:
                    backend = new PstStoreSessionBackend(stream, format, options, cancellationToken);
                    break;
                case EmailStoreFormat.Olm:
                    backend = new MaterializedEmailStoreSessionBackend(
                        new OlmStoreReader(options).Read(stream, sourceName, cancellationToken));
                    break;
                case EmailStoreFormat.Emlx:
                    backend = new MaterializedEmailStoreSessionBackend(
                        new EmlxStoreReader(options).Read(stream, sourceName, cancellationToken));
                    break;
                case EmailStoreFormat.Mbox:
                    backend = new MboxStoreSessionBackend(
                        stream, sourceName, options, cancellationToken);
                    break;
                default:
                    throw new InvalidDataException("The source is not a supported email-store artifact.");
            }
            return new EmailStoreSession(stream, leaveOpen, originalPosition, options, backend);
        } catch {
            if (leaveOpen) {
                if (stream.CanSeek) stream.Position = originalPosition;
            } else {
                stream.Dispose();
            }
            throw;
        }
    }

    private void ThrowIfDisposed() {
        if (_disposed) throw new ObjectDisposedException(nameof(EmailStoreSession));
    }

    private static bool Matches(EmailStoreQuery query, EmailStoreItemSummary summary) {
        if (query.ItemKind.HasValue && summary.OutlookItemKind != query.ItemKind.Value) return false;
        if (query.SubjectContains != null && !Contains(summary.Subject, query.SubjectContains)) return false;
        if (query.SenderContains != null &&
            !AddressContains(summary.From, query.SenderContains) &&
            !AddressContains(summary.Sender, query.SenderContains)) return false;
        DateTimeOffset? timestamp = summary.ReceivedAt ?? summary.SentAt;
        if (query.Since.HasValue && (!timestamp.HasValue || timestamp.Value < query.Since.Value)) return false;
        if (query.Before.HasValue && (!timestamp.HasValue || timestamp.Value >= query.Before.Value)) return false;
        if (query.HasAttachments.HasValue && summary.HasAttachments != query.HasAttachments) return false;
        if (query.IsRead.HasValue && summary.IsRead != query.IsRead) return false;
        return true;
    }

    private static bool AddressContains(OfficeIMO.Email.EmailAddress? address, string value) =>
        address != null && (Contains(address.Address, value) ||
            Contains(address.DisplayName, value) || Contains(address.RawValue, value));

    private static bool Contains(string? text, string value) =>
        text != null && text.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;

}
