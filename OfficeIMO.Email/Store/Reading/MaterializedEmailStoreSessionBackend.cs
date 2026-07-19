namespace OfficeIMO.Email.Store;

internal sealed class MaterializedEmailStoreSessionBackend : IEmailStoreSessionBackend {
    private readonly EmailStoreReadResult _result;
    private readonly IReadOnlyList<EmailStoreFolderInfo> _folders;
    private readonly Dictionary<string, EmailStoreItem> _items;

    internal MaterializedEmailStoreSessionBackend(EmailStoreReadResult result) {
        _result = result ?? throw new ArgumentNullException(nameof(result));
        _folders = result.Store.Folders.Select(folder => new EmailStoreFolderInfo(
            folder.Id, folder.ParentId, folder.Name, folder.Items.Count, folder.AssociatedItems.Count,
            folder.SpecialFolderKind, folder.ClassificationSource,
            folder.ContainerClass, folder.IsSearchFolder, folder.MapiProperties)).ToArray();
        _items = result.Store.Folders
            .SelectMany(folder => folder.Items.Concat(folder.AssociatedItems))
            .ToDictionary(item => item.Id, StringComparer.Ordinal);
    }

    public EmailStoreFormat Format => _result.Store.Format;
    public string? DisplayName => _result.Store.DisplayName;
    public long SourceLength => _result.BytesRead;
    public IReadOnlyList<EmailStoreFolderInfo> Folders => _folders;
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _result.Diagnostics;

    public IEnumerable<EmailStoreItemReference> EnumerateItems(
        EmailStoreEnumerationOptions options, CancellationToken cancellationToken) {
        HashSet<string>? folderIds = ResolveFolderIds(options);
        int count = 0;
        foreach (EmailStoreFolder folder in _result.Store.Folders) {
            cancellationToken.ThrowIfCancellationRequested();
            if (folderIds != null && !folderIds.Contains(folder.Id)) continue;
            if (options.IncludeRegularItems) {
                foreach (EmailStoreItem item in folder.Items) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (item.IsOrphaned && !options.IncludeOrphanedItems) continue;
                    if (++count > options.MaxItems) yield break;
                    yield return new EmailStoreItemReference(
                        item.Id, folder.Id, false, item.IsOrphaned, EmailStoreItemSummary.FromItem(item));
                }
            }
            if (!options.IncludeAssociatedItems) continue;
            foreach (EmailStoreItem item in folder.AssociatedItems) {
                cancellationToken.ThrowIfCancellationRequested();
                if (item.IsOrphaned && !options.IncludeOrphanedItems) continue;
                if (++count > options.MaxItems) yield break;
                yield return new EmailStoreItemReference(
                    item.Id, folder.Id, true, item.IsOrphaned, EmailStoreItemSummary.FromItem(item));
            }
        }
    }

    public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference,
        CancellationToken cancellationToken) =>
        EmailStoreItemSummary.FromItem(ReadItem(reference, EmailStoreItemReadOptions.Default, cancellationToken));

    public EmailStoreItem ReadItem(EmailStoreItemReference reference, EmailStoreItemReadOptions options,
        CancellationToken cancellationToken) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        if (!_items.TryGetValue(reference.Id, out EmailStoreItem? item) || item.FolderId != reference.FolderId) {
            throw new KeyNotFoundException("The item reference does not belong to this email-store session.");
        }
        return item;
    }

    public void Dispose() { }

    private HashSet<string>? ResolveFolderIds(EmailStoreEnumerationOptions options) {
        if (options.FolderId == null) return null;
        if (!_folders.Any(folder => folder.Id == options.FolderId)) {
            throw new KeyNotFoundException("The requested folder does not belong to this email-store session.");
        }
        var result = new HashSet<string>(StringComparer.Ordinal) { options.FolderId };
        if (!options.IncludeDescendants) return result;
        bool added;
        do {
            added = false;
            foreach (EmailStoreFolderInfo folder in _folders) {
                if (folder.ParentId != null && result.Contains(folder.ParentId) && result.Add(folder.Id)) added = true;
            }
        } while (added);
        return result;
    }
}
