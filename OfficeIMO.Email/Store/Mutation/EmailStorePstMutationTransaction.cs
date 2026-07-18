using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>
/// Stages folder and item changes against an existing Unicode PST and commits them as a verified atomic rewrite.
/// Disposing an uncommitted transaction leaves the source byte-for-byte unchanged.
/// </summary>
public sealed partial class EmailStorePstMutationTransaction : IDisposable {
    private readonly string _sourcePath;
    private readonly EmailStorePstMutationOptions _options;
    private readonly long _sourceLength;
    private readonly DateTime _sourceLastWriteTimeUtc;
    private readonly Dictionary<string, FolderState> _folders;
    private readonly List<EmailStoreDiagnostic> _diagnostics;
    private PstMutationTransactionLock? _transactionLock;
    private EmailStoreSession? _source;
    private Dictionary<string, ItemState>? _items;
    private bool _committed;
    private bool _disposed;

    private EmailStorePstMutationTransaction(string sourcePath,
        EmailStorePstMutationOptions options, EmailStoreSession source,
        FileInfo sourceFile, PstMutationTransactionLock transactionLock) {
        _sourcePath = sourcePath;
        _options = options;
        _source = source;
        _sourceLength = sourceFile.Length;
        _sourceLastWriteTimeUtc = sourceFile.LastWriteTimeUtc;
        _folders = source.Folders.ToDictionary(folder => folder.Id,
            folder => new FolderState(folder, source.IsOfficeImoWriterStore), StringComparer.Ordinal);
        _diagnostics = new List<EmailStoreDiagnostic>(source.Diagnostics);
        _transactionLock = transactionLock;
        RootFolderId = ResolveRootFolderId(source.Folders);
    }

    /// <summary>Opens and locks an existing Unicode PST for one mutation transaction.</summary>
    public static EmailStorePstMutationTransaction Open(string path,
        EmailStorePstMutationOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("An existing PST path is required.", nameof(path));
        }
        string sourcePath = EmailStorePathIdentity.ResolvePhysicalPath(path);
        if (!File.Exists(sourcePath)) throw new FileNotFoundException("The PST does not exist.", sourcePath);
        var effective = options ?? new EmailStorePstMutationOptions();
        if (effective.BackupPath != null && EmailStorePathIdentity.AreEquivalent(
            sourcePath, effective.BackupPath)) {
            throw new ArgumentException("The backup path must differ from the source PST.", nameof(options));
        }
        if (effective.BackupPath != null && File.Exists(effective.BackupPath) &&
            !effective.OverwriteBackup) {
            throw new IOException("The backup path already exists and overwriteBackup is false.");
        }
        var readerOptions = new EmailStoreReaderOptions(
            maxFolderCount: effective.MaxFolderCount,
            maxItemCount: effective.MaxItemCount,
            retainAttachmentContent: false,
            pstPassword: effective.PstPassword,
            pstPasswordEncoding: effective.PstPasswordEncoding,
            includeAssociatedItems: true,
            includeOrphanedItems: true,
            maxNestedMessageDepth: effective.MaxNestedMessageDepth);
        EmailStoreSession? source = null;
        FileStream? input = null;
        PstMutationTransactionLock? transactionLock = null;
        try {
            transactionLock = PstMutationTransactionLock.Acquire(sourcePath);
            input = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read,
                64 * 1024, FileOptions.RandomAccess);
            PstHeader header = PstHeader.Read(input, EmailStoreFormat.Pst);
            if (!header.IsUnicode) {
                throw new NotSupportedException(
                    "Existing-store mutation supports Unicode PST files only; ANSI PST files remain read-only sources.");
            }
            input.Position = 0;
            source = EmailStoreSession.Open(input, Path.GetFileName(sourcePath),
                readerOptions, leaveOpen: false, cancellationToken);
            input = null;
            if (source.Format != EmailStoreFormat.Pst) {
                throw new NotSupportedException(
                    "Existing-store mutation supports Unicode PST files only; OST files remain read-only sources.");
            }
            if (source.IsPstPasswordProtected) {
                throw new NotSupportedException(
                    "Existing-store mutation rejects password-protected PST files because the managed writer cannot preserve password protection.");
            }
            if (source.Folders.Count > effective.MaxFolderCount) {
                throw new EmailStoreLimitExceededException(
                    nameof(EmailStorePstMutationOptions.MaxFolderCount),
                    source.Folders.Count, effective.MaxFolderCount);
            }
            var sourceFile = new FileInfo(sourcePath);
            var transaction = new EmailStorePstMutationTransaction(
                sourcePath, effective, source, sourceFile, transactionLock);
            transactionLock = null;
            return transaction;
        } catch {
            source?.Dispose();
            input?.Dispose();
            transactionLock?.Dispose();
            throw;
        }
    }

    /// <summary>Full path of the PST that will be replaced only after a successful commit.</summary>
    public string SourcePath { get { ThrowIfUnavailable(); return _sourcePath; } }

    /// <summary>Source folder identifier used when a new folder omits its parent.</summary>
    public string RootFolderId { get; }

    /// <summary>Current staged folder view. Transaction-local identifiers remain stable until commit.</summary>
    public IReadOnlyList<EmailStorePstMutationFolder> Folders {
        get {
            ThrowIfUnavailable();
            return _folders.Values.Where(folder => !folder.Deleted)
                .Select(folder => folder.ToPublic())
                .OrderBy(folder => folder.Id, StringComparer.Ordinal)
                .ToArray();
        }
    }

    /// <summary>Enumerates active source item references after staged deletes.</summary>
    public IEnumerable<EmailStoreItemReference> EnumerateItems() {
        ThrowIfUnavailable();
        EnsureItemsIndexed(default);
        foreach (ItemState item in _items!.Values) {
            ThrowIfUnavailable();
            if (item.Source != null && !item.Deleted) yield return item.Source;
        }
    }

    /// <summary>Reads one source item by its stable source identifier.</summary>
    public EmailStoreItem ReadItem(string itemId, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(itemId)) throw new ArgumentException("An item identifier is required.", nameof(itemId));
        ThrowIfUnavailable();
        ItemState item = GetItem(itemId, cancellationToken);
        if (item.Source == null) {
            throw new InvalidOperationException("A newly staged item has no source-store representation.");
        }
        return _source!.ReadItem(item.Source, cancellationToken);
    }

    /// <inheritdoc />
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _source?.Dispose();
        _source = null;
        _transactionLock?.Dispose();
        _transactionLock = null;
    }

    private static string ResolveRootFolderId(IReadOnlyList<EmailStoreFolderInfo> folders) {
        EmailStoreFolderInfo? root = folders.FirstOrDefault(folder =>
            folder.SpecialFolderKind == EmailStoreSpecialFolderKind.IpmSubtree) ??
            folders.FirstOrDefault(folder =>
                folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Root) ??
            folders.FirstOrDefault(folder => folder.ParentId == null);
        if (root == null) throw new InvalidDataException("The PST does not expose a root folder.");
        return root.Id;
    }

    private void EnsureItemsIndexed(CancellationToken cancellationToken) {
        if (_items != null) return;
        var items = new Dictionary<string, ItemState>(StringComparer.Ordinal);
        int maximum = _options.MaxItemCount == int.MaxValue
            ? int.MaxValue
            : _options.MaxItemCount + 1;
        var enumeration = new EmailStoreEnumerationOptions(
            includeAssociatedItems: true,
            includeOrphanedItems: true,
            maxItems: maximum);
        foreach (EmailStoreItemReference reference in _source!.EnumerateItems(
            enumeration, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (items.Count >= _options.MaxItemCount) {
                throw new EmailStoreLimitExceededException(
                    nameof(EmailStorePstMutationOptions.MaxItemCount),
                    items.Count + 1L, _options.MaxItemCount);
            }
            if (!_folders.ContainsKey(reference.FolderId)) {
                throw new InvalidDataException("A source PST item references an unknown folder.");
            }
            if (items.ContainsKey(reference.Id)) {
                throw new InvalidDataException("The PST exposed a duplicate item identifier.");
            }
            items.Add(reference.Id, new ItemState(reference));
        }
        _items = items;
    }

    private FolderState GetFolder(string folderId) {
        if (string.IsNullOrWhiteSpace(folderId)) {
            throw new ArgumentException("A folder identifier is required.", nameof(folderId));
        }
        if (!_folders.TryGetValue(folderId, out FolderState? folder) || folder.Deleted) {
            throw new KeyNotFoundException("The folder does not exist in the active transaction.");
        }
        return folder;
    }

    private ItemState GetItem(string itemId, CancellationToken cancellationToken) {
        EnsureItemsIndexed(cancellationToken);
        if (!_items!.TryGetValue(itemId, out ItemState? item) || item.Deleted) {
            throw new KeyNotFoundException("The item does not exist in the active transaction.");
        }
        return item;
    }

    private bool HasChanges() {
        if (_folders.Values.Any(folder => folder.HasChanges)) return true;
        return _items != null && _items.Values.Any(item => item.HasChanges);
    }

    private void ThrowIfUnavailable() {
        if (_disposed) throw new ObjectDisposedException(nameof(EmailStorePstMutationTransaction));
        if (_committed) throw new InvalidOperationException("The PST mutation transaction has already committed.");
    }

    private sealed class FolderState {
        internal FolderState(EmailStoreFolderInfo folder, bool isOfficeImoWriterStore) {
            Id = folder.Id;
            SourceId = folder.Id;
            ParentId = folder.ParentId;
            OriginalParentId = folder.ParentId;
            Name = folder.Name;
            OriginalName = folder.Name;
            ContainerClass = folder.ContainerClass;
            SpecialFolderKind = folder.SpecialFolderKind;
            ClassificationSource = folder.ClassificationSource;
            IsSearchFolder = folder.IsSearchFolder;
            IsMappedSystemFolder = isOfficeImoWriterStore && folder.IsSearchFolder &&
                PstStoreWriterCore.IsWriterOwnedSearchFolderId(folder.Id);
        }

        internal FolderState(string id, string parentId, string name, string? containerClass) {
            Id = id;
            ParentId = parentId;
            Name = name;
            OriginalName = name;
            ContainerClass = containerClass;
        }

        internal string Id { get; }
        internal string? SourceId { get; }
        internal string? ParentId { get; set; }
        internal string? OriginalParentId { get; }
        internal string Name { get; set; }
        internal string OriginalName { get; }
        internal string? ContainerClass { get; }
        internal EmailStoreSpecialFolderKind SpecialFolderKind { get; }
        internal EmailStoreFolderClassificationSource ClassificationSource { get; }
        internal bool IsSearchFolder { get; }
        internal bool IsMappedSystemFolder { get; set; }
        internal bool Deleted { get; set; }
        internal bool IsCreated => SourceId == null;
        internal bool IsMandatory => IsMappedSystemFolder ||
            ClassificationSource == EmailStoreFolderClassificationSource.SourceIdentifier &&
            (SpecialFolderKind == EmailStoreSpecialFolderKind.Root ||
             SpecialFolderKind == EmailStoreSpecialFolderKind.IpmSubtree ||
             SpecialFolderKind == EmailStoreSpecialFolderKind.DeletedItems ||
             SpecialFolderKind == EmailStoreSpecialFolderKind.SearchRoot);
        internal bool HasChanges => IsCreated ? !Deleted : Deleted ||
            !string.Equals(Name, OriginalName, StringComparison.Ordinal) ||
            !string.Equals(ParentId, OriginalParentId, StringComparison.Ordinal);

        internal EmailStorePstMutationFolder ToPublic() => new EmailStorePstMutationFolder(
            Id, ParentId, Name, ContainerClass, SpecialFolderKind, IsSearchFolder, IsCreated);
    }

    private sealed class ItemState {
        internal ItemState(EmailStoreItemReference source) {
            Id = source.Id;
            Source = source;
            FolderId = source.FolderId;
            OriginalFolderId = source.FolderId;
            IsAssociated = source.IsAssociated;
            OriginalIsAssociated = source.IsAssociated;
        }

        internal ItemState(string id, string folderId, EmailDocument document, bool isAssociated) {
            Id = id;
            FolderId = folderId;
            OriginalFolderId = folderId;
            Document = document;
            IsAssociated = isAssociated;
            OriginalIsAssociated = isAssociated;
        }

        internal string Id { get; }
        internal EmailStoreItemReference? Source { get; }
        internal string FolderId { get; set; }
        internal string OriginalFolderId { get; }
        internal EmailDocument? Document { get; set; }
        internal bool IsAssociated { get; set; }
        internal bool OriginalIsAssociated { get; }
        internal bool Replaced { get; set; }
        internal int PropertyPatchChanges { get; set; }
        internal int AttachmentPatchChanges { get; set; }
        internal string? CopiedFromId { get; set; }
        internal bool Deleted { get; set; }
        internal bool IsCreated => Source == null;
        internal bool HasChanges => IsCreated ? !Deleted : Deleted || Replaced ||
            PropertyPatchChanges > 0 || AttachmentPatchChanges > 0 ||
            !string.Equals(FolderId, OriginalFolderId, StringComparison.Ordinal) ||
            IsAssociated != OriginalIsAssociated;
    }
}

/// <summary>One folder in the transaction's current staged hierarchy.</summary>
public sealed class EmailStorePstMutationFolder {
    internal EmailStorePstMutationFolder(string id, string? parentId, string name,
        string? containerClass, EmailStoreSpecialFolderKind specialFolderKind,
        bool isSearchFolder, bool isCreated) {
        Id = id;
        ParentId = parentId;
        Name = name;
        ContainerClass = containerClass;
        SpecialFolderKind = specialFolderKind;
        IsSearchFolder = isSearchFolder;
        IsCreated = isCreated;
    }

    /// <summary>Stable source or transaction-local identifier.</summary>
    public string Id { get; }

    /// <summary>Current staged parent identifier.</summary>
    public string? ParentId { get; }

    /// <summary>Current staged display name.</summary>
    public string Name { get; }

    /// <summary>MAPI container class.</summary>
    public string? ContainerClass { get; }

    /// <summary>Well-known folder role inherited from the source.</summary>
    public EmailStoreSpecialFolderKind SpecialFolderKind { get; }

    /// <summary>Whether the source identifies this as a search folder.</summary>
    public bool IsSearchFolder { get; }

    /// <summary>Whether the folder was created by this transaction.</summary>
    public bool IsCreated { get; }
}
