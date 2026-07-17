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
        while (ancestor != null) {
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
        var removed = new HashSet<string>(StringComparer.Ordinal) { folder.Id };
        bool changed;
        do {
            changed = false;
            foreach (FolderState candidate in _folders.Values) {
                if (candidate.Deleted || candidate.ParentId == null || removed.Contains(candidate.Id)) continue;
                if (removed.Contains(candidate.ParentId) && removed.Add(candidate.Id)) changed = true;
            }
        } while (changed);

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
}
