using System.IO.Compression;

namespace OfficeIMO.Email.Store;

/// <summary>Keeps an OLM catalog and projects selected XML records without retaining every document.</summary>
internal sealed class OlmStoreSessionBackend : IEmailStoreSessionBackend {
    private readonly Stream _source;
    private readonly EmailStoreSourceGuard _sourceGuard;
    private readonly ZipArchive _archive;
    private readonly OlmStoreReader _reader;
    private readonly EmailStoreReadResources _resources;
    private readonly List<IndexedItem> _items = new List<IndexedItem>();
    private readonly Dictionary<string, IndexedItem> _byId = new Dictionary<string, IndexedItem>(StringComparer.Ordinal);

    internal OlmStoreSessionBackend(Stream source, string? sourceName, EmailStoreReaderOptions options,
        CancellationToken cancellationToken, bool isSnapshot = false) {
        _source = source;
        SourceLength = source.Length;
        _resources = new EmailStoreReadResources(options);
        _sourceGuard = new EmailStoreSourceGuard(source, SourceLength, options.MaxInputBytes, _resources.Dispose, cancellationToken, isSnapshot);
        _reader = new OlmStoreReader(options);
        source.Position = 0;
        _archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        try {
            EmailStoreReadResult catalog = _reader.ReadArchive(_archive, sourceName, SourceLength, cancellationToken,
                (item, path, index, kind) => {
                    var reference = new EmailStoreItemReference(item.Id, item.FolderId, false, false,
                        EmailStoreItemSummary.FromItem(item));
                    var indexed = new IndexedItem(reference, path, index, kind);
                    _items.Add(indexed);
                    _byId.Add(reference.Id, indexed);
                });
            DisplayName = catalog.Store.DisplayName;
            var counts = _items.GroupBy(item => item.Reference.FolderId)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
            Folders = catalog.Store.Folders.Select(folder => new EmailStoreFolderInfo(folder.Id,
                folder.ParentId, folder.Name, counts.TryGetValue(folder.Id, out int count) ? count : 0, 0)).ToArray();
            _sourceGuard.Validate(source, cancellationToken);
        } catch {
            _archive.Dispose();
            _resources.Dispose();
            throw;
        }
    }

    public EmailStoreFormat Format => EmailStoreFormat.Olm;
    public string? DisplayName { get; }
    public long SourceLength { get; }
    public IReadOnlyList<EmailStoreFolderInfo> Folders { get; }
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _reader.Diagnostics;

    public IEnumerable<EmailStoreItemReference> EnumerateItems(EmailStoreEnumerationOptions options,
        CancellationToken cancellationToken) {
        HashSet<string>? scope = null;
        if (options.FolderId != null) {
            if (!Folders.Any(folder => folder.Id == options.FolderId))
                throw new KeyNotFoundException("The folder does not belong to this OLM session.");
            scope = new HashSet<string>(StringComparer.Ordinal) { options.FolderId };
            if (options.IncludeDescendants) {
                // Folder catalogs are created in parent-before-child order.
                foreach (EmailStoreFolderInfo folder in Folders)
                    if (folder.ParentId != null && scope.Contains(folder.ParentId)) scope.Add(folder.Id);
            }
        }
        if (!options.IncludeRegularItems) yield break;
        int count = 0;
        foreach (IndexedItem item in _items) {
            cancellationToken.ThrowIfCancellationRequested();
            if (scope != null && !scope.Contains(item.Reference.FolderId)) continue;
            if (++count > options.MaxItems) yield break;
            yield return item.Reference;
        }
    }

    public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference, CancellationToken cancellationToken) =>
        GetItem(reference, cancellationToken).Reference.Summary!;

    public EmailStoreItem ReadItem(EmailStoreItemReference reference, EmailStoreItemReadOptions options,
        CancellationToken cancellationToken) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        IndexedItem item = GetItem(reference, cancellationToken);
        EmailReadWorkspace? workspace = options.PreferStreamingAttachmentContent &&
            options.Includes(EmailStoreItemReadParts.AttachmentContent) ? new EmailReadWorkspace() : null;
        try {
            EmailStoreItem result = _reader.ReadSelected(item.Path, item.Index, item.Kind, reference.Id,
                reference.FolderId, options, workspace, cancellationToken);
            _sourceGuard.Validate(_source, cancellationToken);
            if (workspace != null) {
                using (var content = new EmailReadResult(result.Document, Array.Empty<EmailDiagnostic>(), 0, workspace)) {
                    _resources.Adopt(content);
                }
                workspace = null;
            }
            return result;
        } finally {
            workspace?.Dispose();
        }
    }

    private IndexedItem GetItem(EmailStoreItemReference reference, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!_byId.TryGetValue(reference.Id, out IndexedItem? item) ||
            item.Reference.FolderId != reference.FolderId || reference.IsAssociated || reference.IsOrphaned)
            throw new KeyNotFoundException("The item reference does not belong to this OLM session.");
        _sourceGuard.Validate(_source, cancellationToken);
        return item;
    }

    public void Dispose() {
        _resources.Dispose();
        _archive.Dispose();
    }

    private sealed class IndexedItem {
        internal IndexedItem(EmailStoreItemReference reference, string path, int index, OutlookItemKind kind) {
            Reference = reference; Path = path; Index = index; Kind = kind;
        }
        internal EmailStoreItemReference Reference { get; }
        internal string Path { get; }
        internal int Index { get; }
        internal OutlookItemKind Kind { get; }
    }
}
