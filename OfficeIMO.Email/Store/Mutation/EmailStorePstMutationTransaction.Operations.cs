using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStorePstMutationTransaction {
    /// <summary>Creates a folder and returns its transaction-local identifier.</summary>
    public string CreateFolder(string name, string? parentFolderId = null,
        string? containerClass = null) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("A folder name is required.", nameof(name));
        }
        ThrowIfUnavailable();
        string parent = parentFolderId ?? RootFolderId;
        FolderState parentFolder = GetFolder(parent);
        if (parentFolder.IsSearchFolder) {
            throw new InvalidOperationException("A normal folder cannot be created below a search folder.");
        }
        int activeFolders = _folders.Values.Count(folder => !folder.Deleted);
        if (activeFolders >= _options.MaxFolderCount) {
            throw new EmailStoreLimitExceededException(
                nameof(EmailStorePstMutationOptions.MaxFolderCount),
                activeFolders + 1L, _options.MaxFolderCount);
        }
        string id;
        do {
            id = string.Concat("mutation-folder:", Guid.NewGuid().ToString("N"));
        } while (_folders.ContainsKey(id));
        _folders.Add(id, new FolderState(id, parent, name, containerClass));
        return id;
    }

    /// <summary>Changes a non-mandatory folder's display name.</summary>
    public void RenameFolder(string folderId, string name) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("A folder name is required.", nameof(name));
        }
        ThrowIfUnavailable();
        FolderState folder = GetFolder(folderId);
        if (folder.IsMandatory) {
            throw new InvalidOperationException("Mandatory PST folders cannot be renamed by a rewrite transaction.");
        }
        folder.Name = name;
    }

    /// <summary>Moves a non-mandatory folder below another active folder.</summary>
    public void MoveFolder(string folderId, string newParentFolderId) {
        ThrowIfUnavailable();
        FolderState folder = GetFolder(folderId);
        FolderState parent = GetFolder(newParentFolderId);
        if (folder.IsMandatory) {
            throw new InvalidOperationException("Mandatory PST folders cannot be moved by a rewrite transaction.");
        }
        if (parent.IsSearchFolder) {
            throw new InvalidOperationException("A normal folder cannot be moved below a search folder.");
        }
        if (ReferenceEquals(folder, parent)) {
            throw new InvalidOperationException("A folder cannot be its own parent.");
        }
        FolderState? ancestor = parent;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (ancestor != null) {
            if (!visited.Add(ancestor.Id)) {
                throw new InvalidOperationException(
                    "The source folder hierarchy already contains a cycle; the move was not staged.");
            }
            if (ReferenceEquals(ancestor, folder)) {
                throw new InvalidOperationException("Moving the folder would create a hierarchy cycle.");
            }
            ancestor = ancestor.ParentId != null && _folders.TryGetValue(
                ancestor.ParentId, out FolderState? next) && !next.Deleted
                    ? next
                    : null;
        }
        folder.ParentId = parent.Id;
    }

    /// <summary>
    /// Deletes a non-mandatory folder. A non-empty folder requires <paramref name="recursive"/>.
    /// Recursive deletion includes active descendants and their items.
    /// </summary>
    public void DeleteFolder(string folderId, bool recursive = false,
        CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        FolderState folder = GetFolder(folderId);
        if (folder.IsMandatory) {
            throw new InvalidOperationException("Mandatory PST folders cannot be deleted by a rewrite transaction.");
        }
        EnsureItemsIndexed(cancellationToken);
        Dictionary<string, FolderState[]> childrenByParent = _folders.Values
            .Where(candidate => !candidate.Deleted && candidate.ParentId != null)
            .GroupBy(candidate => candidate.ParentId!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        var removed = new HashSet<string>(StringComparer.Ordinal);
        var pending = new Queue<FolderState>();
        pending.Enqueue(folder);
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            FolderState current = pending.Dequeue();
            if (current.IsMandatory) {
                throw new InvalidOperationException(
                    "Recursive deletion cannot include a mandatory PST folder, even when its source parent relationship is malformed.");
            }
            if (!removed.Add(current.Id) ||
                !childrenByParent.TryGetValue(current.Id, out FolderState[]? children)) continue;
            foreach (FolderState child in children) pending.Enqueue(child);
        }

        bool hasChildren = removed.Count > 1;
        bool hasItems = _items!.Values.Any(item => !item.Deleted && removed.Contains(item.FolderId));
        if (!recursive && (hasChildren || hasItems)) {
            throw new InvalidOperationException("The folder is not empty; set recursive to delete its descendants and items.");
        }
        foreach (string id in removed) _folders[id].Deleted = true;
        foreach (ItemState item in _items.Values) {
            if (!item.Deleted && removed.Contains(item.FolderId)) item.Deleted = true;
        }
    }

    /// <summary>Adds an item and returns its transaction-local identifier.</summary>
    public string AddItem(string folderId, EmailDocument document, bool isAssociated = false,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        ThrowIfUnavailable();
        EnsureWritableItemFolder(GetFolder(folderId));
        EnsureItemsIndexed(cancellationToken);
        int activeItems = _items!.Values.Count(item => !item.Deleted);
        if (activeItems >= _options.MaxItemCount) {
            throw new EmailStoreLimitExceededException(
                nameof(EmailStorePstMutationOptions.MaxItemCount),
                activeItems + 1L, _options.MaxItemCount);
        }
        string id;
        do {
            id = string.Concat("mutation-item:", Guid.NewGuid().ToString("N"));
        } while (_items.ContainsKey(id));
        _items.Add(id, new ItemState(id, folderId, document, isAssociated));
        return id;
    }

    /// <summary>Replaces the semantic document for an existing or newly staged item.</summary>
    public void ReplaceItem(string itemId, EmailDocument document,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        ThrowIfUnavailable();
        ItemState item = GetItem(itemId, cancellationToken);
        item.Document = document;
        if (!item.IsCreated) item.Replaced = true;
    }

    /// <summary>Copies an existing source item into another folder and returns a transaction-local identifier.</summary>
    public string CopyItem(string itemId, string destinationFolderId, bool? isAssociated = null,
        CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        EnsureWritableItemFolder(GetFolder(destinationFolderId));
        ItemState source = GetItem(itemId, cancellationToken);
        if (source.IsCreated) {
            throw new InvalidOperationException(
                "CopyItem requires an existing source-store item; add the staged document again when copying a newly staged item.");
        }
        var readOptions = new EmailStoreItemReadOptions(
            EmailStoreItemReadParts.All, preferStreamingAttachmentContent: true);
        EmailDocument copy = _source!.ReadItem(source.Source!, readOptions, cancellationToken).Document;
        string copyId = AddItem(destinationFolderId, copy,
            isAssociated ?? source.IsAssociated, cancellationToken);
        _items![copyId].CopiedFromId = source.Id;
        return copyId;
    }

    /// <summary>Copies a source folder, optionally including its complete subtree and items.</summary>
    public EmailStorePstFolderCopyResult CopyFolder(string folderId, string destinationParentFolderId,
        bool includeDescendants = true,
        EmailStorePstCopyConflictPolicy conflictPolicy = EmailStorePstCopyConflictPolicy.Fail,
        CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        FolderState sourceRoot = GetFolder(folderId);
        FolderState destinationParent = GetFolder(destinationParentFolderId);
        if (destinationParent.IsSearchFolder)
            throw new InvalidOperationException("A folder cannot be copied below a search folder.");
        if (!Enum.IsDefined(typeof(EmailStorePstCopyConflictPolicy), conflictPolicy))
            throw new ArgumentOutOfRangeException(nameof(conflictPolicy));
        if (conflictPolicy == EmailStorePstCopyConflictPolicy.Fail && _folders.Values.Any(folder =>
            !folder.Deleted && string.Equals(folder.ParentId, destinationParent.Id, StringComparison.Ordinal) &&
            string.Equals(folder.Name, sourceRoot.Name, StringComparison.OrdinalIgnoreCase))) {
            throw new InvalidOperationException("The destination already contains a folder with the copied display name.");
        }

        EnsureItemsIndexed(cancellationToken);
        var sourceFolders = new List<FolderState> { sourceRoot };
        if (includeDescendants) {
            var pending = new Queue<FolderState>();
            pending.Enqueue(sourceRoot);
            while (pending.Count > 0) {
                cancellationToken.ThrowIfCancellationRequested();
                FolderState parent = pending.Dequeue();
                foreach (FolderState child in _folders.Values.Where(folder => !folder.Deleted &&
                    string.Equals(folder.ParentId, parent.Id, StringComparison.Ordinal))
                    .OrderBy(folder => folder.Id, StringComparer.Ordinal)) {
                    sourceFolders.Add(child);
                    pending.Enqueue(child);
                }
            }
        }
        int prospectiveFolders = _folders.Values.Count(folder => !folder.Deleted) + sourceFolders.Count;
        int sourceItemCount = _items!.Values.Count(item => !item.Deleted && !item.IsCreated &&
            sourceFolders.Any(folder => string.Equals(folder.Id, item.FolderId, StringComparison.Ordinal)));
        int prospectiveItems = _items.Values.Count(item => !item.Deleted) + sourceItemCount;
        if (prospectiveFolders > _options.MaxFolderCount)
            throw new EmailStoreLimitExceededException(nameof(EmailStorePstMutationOptions.MaxFolderCount),
                prospectiveFolders, _options.MaxFolderCount);
        if (prospectiveItems > _options.MaxItemCount)
            throw new EmailStoreLimitExceededException(nameof(EmailStorePstMutationOptions.MaxItemCount),
                prospectiveItems, _options.MaxItemCount);

        var folderMap = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (FolderState source in sourceFolders) {
            cancellationToken.ThrowIfCancellationRequested();
            string parent = ReferenceEquals(source, sourceRoot)
                ? destinationParent.Id
                : folderMap[source.ParentId!];
            folderMap.Add(source.Id, CreateFolder(source.Name, parent, source.ContainerClass));
        }
        var itemMap = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (ItemState source in _items.Values.Where(item => !item.Deleted && !item.IsCreated &&
            folderMap.ContainsKey(item.FolderId)).OrderBy(item => item.Id, StringComparer.Ordinal).ToArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            itemMap.Add(source.Id, CopyItem(source.Id, folderMap[source.FolderId],
                source.IsAssociated, cancellationToken));
        }
        return new EmailStorePstFolderCopyResult(folderMap[sourceRoot.Id],
            new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(folderMap),
            new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(itemMap));
    }

    /// <summary>Applies one reusable typed patch to an item.</summary>
    public void PatchItem(string itemId, EmailStoreItemPatch patch,
        CancellationToken cancellationToken = default) {
        if (patch == null) throw new ArgumentNullException(nameof(patch));
        ThrowIfUnavailable();
        if (patch.IsEmpty) return;
        ItemState item = GetItem(itemId, cancellationToken);
        EmailDocument document = GetMutableDocument(item, cancellationToken);
        int changes = patch.Apply(document);
        item.PropertyPatchChanges = checked(item.PropertyPatchChanges +
            changes - patch.Attachments.Changes.Count);
        item.AttachmentPatchChanges = checked(item.AttachmentPatchChanges +
            patch.Attachments.Changes.Count);
    }

    /// <summary>
    /// Applies one reusable typed patch to every row selected by a bounded table query. Selection completes before
    /// any item is changed, and a scan-limit result is rejected rather than partially mutating the transaction.
    /// </summary>
    public EmailStorePstMutationSelectionReport PatchItems(EmailStoreTableQuery query,
        EmailStoreItemPatch patch, CancellationToken cancellationToken = default) {
        if (query == null) throw new ArgumentNullException(nameof(query));
        if (patch == null) throw new ArgumentNullException(nameof(patch));
        ThrowIfUnavailable();
        if (query.ContinuationToken != null)
            throw new ArgumentException("A mutation selection query must start without a continuation token.", nameof(query));
        var selected = new List<EmailStoreItemId>();
        EmailStoreTableQuery pageQuery = query;
        int itemsScanned = 0;
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            EmailStoreTablePage page = _source!.SearchPage(pageQuery, cancellationToken);
            itemsScanned = Math.Max(itemsScanned, page.ItemsScanned);
            if (page.ScanLimitReached) {
                throw new InvalidOperationException(
                    "The bounded mutation query reached MaxItemsScanned; no selected item was patched.");
            }
            selected.AddRange(page.Rows.Select(row => row.Reference.Key));
            if (page.NextToken == null) break;
            pageQuery = query.ContinueFrom(page.NextToken);
        }
        var targets = new List<Tuple<ItemState, EmailDocument>>(selected.Count);
        foreach (EmailStoreItemId itemId in selected) {
            cancellationToken.ThrowIfCancellationRequested();
            ItemState item = GetItem(itemId.Value, cancellationToken);
            EmailDocument document = GetMutableDocument(item, cancellationToken);
            patch.Validate(document);
            targets.Add(Tuple.Create(item, document));
        }
        foreach (Tuple<ItemState, EmailDocument> target in targets) {
            int changes = patch.Apply(target.Item2);
            target.Item1.PropertyPatchChanges = checked(target.Item1.PropertyPatchChanges +
                changes - patch.Attachments.Changes.Count);
            target.Item1.AttachmentPatchChanges = checked(target.Item1.AttachmentPatchChanges +
                patch.Attachments.Changes.Count);
        }
        return new EmailStorePstMutationSelectionReport(selected.AsReadOnly(), itemsScanned);
    }

    /// <summary>
    /// Applies exact typed MAPI property changes after normal semantic projection when the item is rewritten.
    /// </summary>
    public void PatchItemProperties(string itemId, MapiPropertyPatch patch,
        CancellationToken cancellationToken = default) {
        if (patch == null) throw new ArgumentNullException(nameof(patch));
        ThrowIfUnavailable();
        if (patch.IsEmpty) return;
        ItemState item = GetItem(itemId, cancellationToken);
        EmailDocument document = GetMutableDocument(item, cancellationToken);
        document.MapiWritePatch.Append(patch);
        item.PropertyPatchChanges = checked(item.PropertyPatchChanges + patch.Changes.Count);
    }

    /// <summary>Applies ordered, bounds-checked attachment changes to an item.</summary>
    public void PatchItemAttachments(string itemId, EmailAttachmentPatch patch,
        CancellationToken cancellationToken = default) {
        if (patch == null) throw new ArgumentNullException(nameof(patch));
        ThrowIfUnavailable();
        if (patch.IsEmpty) return;
        ItemState item = GetItem(itemId, cancellationToken);
        EmailDocument document = GetMutableDocument(item, cancellationToken);
        patch.Apply(document.Attachments);
        item.AttachmentPatchChanges = checked(item.AttachmentPatchChanges + patch.Changes.Count);
    }

    /// <summary>Returns a value-free dry run of every effective staged operation without writing an artifact.</summary>
    public EmailStorePstMutationPlan DryRun(CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        EnsureItemsIndexed(cancellationToken);
        var operations = new List<EmailStorePstMutationPlanOperation>();
        foreach (FolderState folder in _folders.Values.OrderBy(value => value.Id, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (folder.Deleted) {
                if (!folder.IsCreated) operations.Add(new EmailStorePstMutationPlanOperation(
                    EmailStorePstMutationOperationKind.DeleteFolder, folder.Id, null));
                continue;
            }
            if (folder.IsCreated) operations.Add(new EmailStorePstMutationPlanOperation(
                EmailStorePstMutationOperationKind.CreateFolder, folder.Id, folder.ParentId));
            else {
                if (!string.Equals(folder.Name, folder.OriginalName, StringComparison.Ordinal))
                    operations.Add(new EmailStorePstMutationPlanOperation(
                        EmailStorePstMutationOperationKind.RenameFolder, folder.Id, null));
                if (!string.Equals(folder.ParentId, folder.OriginalParentId, StringComparison.Ordinal))
                    operations.Add(new EmailStorePstMutationPlanOperation(
                        EmailStorePstMutationOperationKind.MoveFolder, folder.Id, folder.ParentId));
            }
        }
        foreach (ItemState item in _items!.Values.OrderBy(value => value.Id, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (item.Deleted) {
                if (!item.IsCreated) operations.Add(new EmailStorePstMutationPlanOperation(
                    EmailStorePstMutationOperationKind.DeleteItem, item.Id, null));
                continue;
            }
            if (item.IsCreated) operations.Add(new EmailStorePstMutationPlanOperation(
                item.CopiedFromId == null ? EmailStorePstMutationOperationKind.AddItem :
                    EmailStorePstMutationOperationKind.CopyItem,
                item.Id, item.CopiedFromId ?? item.FolderId));
            else if (item.Replaced) operations.Add(new EmailStorePstMutationPlanOperation(
                EmailStorePstMutationOperationKind.ReplaceItem, item.Id, null));
            int patchChanges = checked(item.PropertyPatchChanges + item.AttachmentPatchChanges);
            if (patchChanges > 0) operations.Add(new EmailStorePstMutationPlanOperation(
                EmailStorePstMutationOperationKind.PatchItem, item.Id, null, patchChanges));
            if (!item.IsCreated && (!string.Equals(item.FolderId, item.OriginalFolderId, StringComparison.Ordinal) ||
                item.IsAssociated != item.OriginalIsAssociated)) {
                operations.Add(new EmailStorePstMutationPlanOperation(
                    EmailStorePstMutationOperationKind.MoveItem, item.Id, item.FolderId));
            }
        }
        return new EmailStorePstMutationPlan(_sourcePath, operations.AsReadOnly(),
            _folders.Values.Count(folder => !folder.Deleted),
            _items.Values.Count(item => !item.Deleted), _sourceLength,
            _diagnostics.ToArray());
    }

    /// <summary>Moves an existing or newly staged item to another active folder.</summary>
    public void MoveItem(string itemId, string destinationFolderId,
        CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        EnsureWritableItemFolder(GetFolder(destinationFolderId));
        ItemState item = GetItem(itemId, cancellationToken);
        item.FolderId = destinationFolderId;
    }

    /// <summary>Moves an item between normal contents and folder-associated contents.</summary>
    public void SetItemAssociated(string itemId, bool isAssociated,
        CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        ItemState item = GetItem(itemId, cancellationToken);
        item.IsAssociated = isAssociated;
    }

    /// <summary>Deletes an existing or newly staged item.</summary>
    public void DeleteItem(string itemId, CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        ItemState item = GetItem(itemId, cancellationToken);
        item.Deleted = true;
    }

    private static void EnsureWritableItemFolder(FolderState folder) {
        if (folder.IsMappedSystemFolder) {
            throw new InvalidOperationException(
                "Items cannot be added or moved into the writer-owned search folder.");
        }
    }

    private EmailDocument GetMutableDocument(ItemState item, CancellationToken cancellationToken) {
        if (item.Document != null) return item.Document;
        var readOptions = new EmailStoreItemReadOptions(
            EmailStoreItemReadParts.All, preferStreamingAttachmentContent: true);
        item.Document = _source!.ReadItem(item.Source!, readOptions, cancellationToken).Document;
        return item.Document;
    }
}
