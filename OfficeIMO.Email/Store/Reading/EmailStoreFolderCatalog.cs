namespace OfficeIMO.Email.Store;

/// <summary>Indexed, cycle-safe navigation over the lightweight folder catalog of one open store.</summary>
public sealed class EmailStoreFolderCatalog {
    private readonly IReadOnlyList<EmailStoreFolderInfo> _folders;
    private readonly Dictionary<EmailStoreFolderId, EmailStoreFolderInfo> _byId;
    private readonly Dictionary<EmailStoreFolderId, IReadOnlyList<EmailStoreFolderInfo>> _children;
    private readonly IReadOnlyList<EmailStoreFolderInfo> _roots;

    internal EmailStoreFolderCatalog(IReadOnlyList<EmailStoreFolderInfo> folders) {
        _folders = folders ?? throw new ArgumentNullException(nameof(folders));
        _byId = new Dictionary<EmailStoreFolderId, EmailStoreFolderInfo>();
        var children = new Dictionary<EmailStoreFolderId, List<EmailStoreFolderInfo>>();
        var roots = new List<EmailStoreFolderInfo>();

        foreach (EmailStoreFolderInfo folder in folders) {
            if (_byId.ContainsKey(folder.Key)) {
                throw new InvalidDataException(string.Concat("The store exposes duplicate folder identifier '", folder.Id, "'."));
            }
            _byId.Add(folder.Key, folder);
        }

        foreach (EmailStoreFolderInfo folder in folders) {
            if (!folder.ParentKey.HasValue || !_byId.ContainsKey(folder.ParentKey.Value)) {
                roots.Add(folder);
                continue;
            }
            if (!children.TryGetValue(folder.ParentKey.Value, out List<EmailStoreFolderInfo>? list)) {
                list = new List<EmailStoreFolderInfo>();
                children.Add(folder.ParentKey.Value, list);
            }
            list.Add(folder);
        }

        _children = children.ToDictionary(
            pair => pair.Key,
            pair => (IReadOnlyList<EmailStoreFolderInfo>)pair.Value.AsReadOnly());
        _roots = roots.AsReadOnly();
    }

    /// <summary>Every folder in source order.</summary>
    public IReadOnlyList<EmailStoreFolderInfo> All => _folders;

    /// <summary>Folders without a resolvable parent.</summary>
    public IReadOnlyList<EmailStoreFolderInfo> Roots => _roots;

    /// <summary>Gets one folder or throws when the identifier is not in this store.</summary>
    public EmailStoreFolderInfo Get(EmailStoreFolderId id) {
        if (!_byId.TryGetValue(id, out EmailStoreFolderInfo? folder)) {
            throw new KeyNotFoundException(string.Concat("Folder '", id.ToString(), "' does not exist in this store."));
        }
        return folder;
    }

    /// <summary>Attempts to get one folder.</summary>
    public bool TryGet(EmailStoreFolderId id, out EmailStoreFolderInfo? folder) => _byId.TryGetValue(id, out folder);

    /// <summary>Gets the resolvable parent, or null for a root or dangling source parent.</summary>
    public EmailStoreFolderInfo? GetParent(EmailStoreFolderId id) {
        EmailStoreFolderInfo folder = Get(id);
        return folder.ParentKey.HasValue && _byId.TryGetValue(folder.ParentKey.Value, out EmailStoreFolderInfo? parent)
            ? parent
            : null;
    }

    /// <summary>Gets immediate children in source order.</summary>
    public IReadOnlyList<EmailStoreFolderInfo> GetChildren(EmailStoreFolderId id) {
        Get(id);
        return _children.TryGetValue(id, out IReadOnlyList<EmailStoreFolderInfo>? result)
            ? result
            : Array.Empty<EmailStoreFolderInfo>();
    }

    /// <summary>Enumerates descendants breadth-first, bounded and protected from source cycles.</summary>
    public IReadOnlyList<EmailStoreFolderInfo> GetDescendants(EmailStoreFolderId id, int maxFolders = 100_000) {
        if (maxFolders <= 0) throw new ArgumentOutOfRangeException(nameof(maxFolders));
        Get(id);
        var result = new List<EmailStoreFolderInfo>();
        var visited = new HashSet<EmailStoreFolderId> { id };
        var pending = new Queue<EmailStoreFolderInfo>(GetChildren(id));
        while (pending.Count > 0) {
            EmailStoreFolderInfo folder = pending.Dequeue();
            if (!visited.Add(folder.Key)) continue;
            result.Add(folder);
            if (result.Count >= maxFolders) {
                if (pending.Count > 0 || GetChildren(folder.Key).Count > 0) {
                    throw new EmailStoreLimitExceededException(nameof(maxFolders), result.Count + 1L, maxFolders);
                }
                break;
            }
            foreach (EmailStoreFolderInfo child in GetChildren(folder.Key)) pending.Enqueue(child);
        }
        return result.AsReadOnly();
    }

    /// <summary>Gets the root-to-folder path, stopping safely if corrupt parent links form a cycle.</summary>
    public IReadOnlyList<EmailStoreFolderInfo> GetPath(EmailStoreFolderId id, int maxDepth = 1_024) {
        if (maxDepth <= 0) throw new ArgumentOutOfRangeException(nameof(maxDepth));
        var reverse = new List<EmailStoreFolderInfo>();
        var visited = new HashSet<EmailStoreFolderId>();
        EmailStoreFolderInfo? current = Get(id);
        while (current != null) {
            if (!visited.Add(current.Key)) {
                throw new InvalidDataException(string.Concat("Folder parent links contain a cycle at '", current.Id, "'."));
            }
            reverse.Add(current);
            if (reverse.Count > maxDepth) {
                throw new EmailStoreLimitExceededException(nameof(maxDepth), reverse.Count, maxDepth);
            }
            current = current.ParentKey.HasValue && _byId.TryGetValue(current.ParentKey.Value, out EmailStoreFolderInfo? parent)
                ? parent
                : null;
        }
        reverse.Reverse();
        return reverse.AsReadOnly();
    }

    /// <summary>Gets every folder classified with the requested well-known role.</summary>
    public IReadOnlyList<EmailStoreFolderInfo> FindSpecialFolders(EmailStoreSpecialFolderKind kind) =>
        _folders.Where(folder => folder.SpecialFolderKind == kind).ToArray();

    /// <summary>Returns the single matching well-known folder, or null; throws if the role is ambiguous.</summary>
    public EmailStoreFolderInfo? FindSpecialFolder(EmailStoreSpecialFolderKind kind) {
        EmailStoreFolderInfo[] matches = _folders.Where(folder => folder.SpecialFolderKind == kind).Take(2).ToArray();
        if (matches.Length > 1) {
            throw new InvalidOperationException(string.Concat("The store exposes more than one folder classified as ", kind.ToString(), "."));
        }
        return matches.Length == 0 ? null : matches[0];
    }
}
