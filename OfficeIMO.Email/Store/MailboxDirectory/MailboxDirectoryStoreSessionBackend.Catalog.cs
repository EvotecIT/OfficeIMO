
namespace OfficeIMO.Email.Store;

internal sealed partial class MailboxDirectoryStoreSessionBackend {
    private readonly HashSet<string> _catalogFolderPaths = new HashSet<string>(StringComparer.Ordinal);

    private void Index(CancellationToken cancellationToken) {
        List<MailboxCandidate> candidates = ScanCatalog(cancellationToken, true, out HashSet<string> folders);
        _catalogFolderPaths.UnionWith(folders);
        var folderCounts = folders.ToDictionary(path => path, _ => 0, StringComparer.Ordinal);
        foreach (MailboxCandidate candidate in candidates.OrderBy(item => item.RelativePath, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            _sourceLength = AddBounded(_sourceLength, new FileInfo(candidate.Path).Length);
            string folderId = GetFolderId(candidate.FolderPath);
            string id = string.Concat("directory:item:", candidate.RelativePath);
            var file = new MailboxFile(id, folderId, candidate.Path, candidate.RelativePath,
                candidate.IsEmlx, candidate.MaildirFlags);
            _files.Add(file);
            if (candidate.IsAttachmentStorage) {
                string directory = Path.GetDirectoryName(candidate.Path)!;
                if (!_partialFiles.TryGetValue(directory, out List<MailboxFile>? parts)) {
                    parts = new List<MailboxFile>();
                    _partialFiles.Add(directory, parts);
                }
                parts.Add(file);
                continue;
            }
            int previousItemCount = _items.Count;
            if (candidate.IsAggregate) {
                var backend = new MboxStoreSessionBackend(() => OpenRegularMailboxFile(candidate.Path),
                    candidate.RelativePath, _options, cancellationToken, _resources);
                foreach (EmailStoreItemReference inner in backend.EnumerateItems(new EmailStoreEnumerationOptions(), cancellationToken)) {
                    string itemId = string.Concat(id, ":", inner.Id);
                    AddItem(new EmailStoreItemReference(itemId, folderId, false, false), cancellationToken);
                    _aggregateItems.Add(itemId, new AggregateItem(backend, inner, folderId, candidate.RelativePath));
                }
                foreach (EmailStoreDiagnostic diagnostic in backend.Diagnostics) AddDiagnostic(diagnostic, candidate.RelativePath);
            } else {
                _filesById.Add(id, file);
                AddItem(new EmailStoreItemReference(id, folderId, false, false), cancellationToken);
            }
            folderCounts.TryGetValue(candidate.FolderPath, out int previous);
            folderCounts[candidate.FolderPath] = previous + _items.Count - previousItemCount;
        }
        foreach (string path in folderCounts.Keys.OrderBy(item => item, StringComparer.Ordinal)) EnsureFolder(path, folderCounts);
    }

    private void AddItem(EmailStoreItemReference item, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_items.Count >= _options.MaxItemCount) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxItemCount),
                _items.Count + 1L, _options.MaxItemCount);
        }
        _items.Add(item);
    }

    private List<MailboxCandidate> ScanCatalog(CancellationToken cancellationToken, bool recordDiagnostics,
        out HashSet<string> folders) {
        var candidates = new List<MailboxCandidate>();
        long aggregateLength = 0;
        long entryCount = 0;
        folders = new HashSet<string>(StringComparer.Ordinal);
        var pending = new Stack<DirectoryCandidate>();
        string root = TrimTrailingDirectorySeparators(_root);
        if (root.EndsWith(".mbox", StringComparison.OrdinalIgnoreCase)) folders.Add(".");
        pending.Push(new DirectoryCandidate(root, 0));
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            DirectoryCandidate current = pending.Pop();
            try {
                // Enumerate lazily: the candidate bound is checked before a whole directory is allocated or sorted.
                foreach (FileSystemInfo entry in new DirectoryInfo(current.Path).EnumerateFileSystemInfos()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (++entryCount > _options.MaxDirectoryEntryCount) {
                        throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDirectoryEntryCount),
                            entryCount, _options.MaxDirectoryEntryCount);
                    }
                    if ((entry.Attributes & FileAttributes.ReparsePoint) != 0) {
                        if (recordDiagnostics) AddDiagnostic(new EmailStoreDiagnostic(
                            "EMAIL_STORE_DIRECTORY_REPARSE_POINT_SKIPPED",
                            "A symbolic link or reparse point was skipped to keep traversal inside the mailbox root.",
                            EmailStoreDiagnosticSeverity.Information, ToRelativePath(entry.FullName)));
                        continue;
                    }
                    if (entry.Name.StartsWith("._", StringComparison.Ordinal)) {
                        if (recordDiagnostics) AddDiagnostic(new EmailStoreDiagnostic(
                            "EMAIL_STORE_DIRECTORY_APPLEDOUBLE_SKIPPED",
                            "An AppleDouble filesystem sidecar was skipped.", EmailStoreDiagnosticSeverity.Information,
                            ToRelativePath(entry.FullName)));
                        continue;
                    }
                    if (entry is DirectoryInfo directory) {
                        if (current.Depth >= _options.MaxDirectoryDepth) {
                            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDirectoryDepth),
                                current.Depth + 1L, _options.MaxDirectoryDepth);
                        }
                        if (!current.IsAttachmentStorage && directory.Name.EndsWith(".mbox", StringComparison.OrdinalIgnoreCase)) {
                            folders.Add(GetLogicalFolderPath(ToRelativePath(directory.FullName)));
                            if (folders.Count > _options.MaxFolderCount) {
                                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxFolderCount),
                                    folders.Count, _options.MaxFolderCount);
                            }
                        }
                        bool attachmentStorage = current.IsAttachmentStorage ||
                            directory.Name == "Attachments" && Directory.Exists(Path.Combine(current.Path, "Messages"));
                        pending.Push(new DirectoryCandidate(directory.FullName, current.Depth + 1, attachmentStorage));
                        continue;
                    }
                    if (!(entry is FileInfo file) || !current.IsAttachmentStorage && !IsMailboxFile(file)) continue;
                    if (!IsRegularMailboxFile(file.FullName)) {
                        if (recordDiagnostics) AddDiagnostic(new EmailStoreDiagnostic(
                            "EMAIL_STORE_DIRECTORY_SPECIAL_FILE_SKIPPED",
                            "A non-regular mailbox candidate was skipped without opening it as a blocking stream.",
                            EmailStoreDiagnosticSeverity.Warning, ToRelativePath(file.FullName)));
                        continue;
                    }
                    if (candidates.Count >= _options.MaxDirectoryFileCount) {
                        throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDirectoryFileCount),
                            candidates.Count + 1L, _options.MaxDirectoryFileCount);
                    }
                    aggregateLength = AddBounded(aggregateLength, file.Length);
                    candidates.Add(new MailboxCandidate(file.FullName, ToRelativePath(file.FullName),
                        IsEmlx(file), GetLogicalFolderPath(ToRelativePath(file.DirectoryName ?? _root)),
                        ParseMaildirFlags(file.Name, file.Directory?.Name), IsAggregateMailbox(file), current.IsAttachmentStorage));
                }
            } catch (Exception exception) when (!(exception is EmailStoreLimitExceededException) &&
                (exception is IOException || exception is UnauthorizedAccessException)) {
                if (!recordDiagnostics) throw new InvalidDataException("The mailbox-directory catalog is no longer readable.", exception);
                AddDiagnostic(new EmailStoreDiagnostic("EMAIL_STORE_DIRECTORY_ENUMERATION_FAILED", exception.Message,
                    EmailStoreDiagnosticSeverity.Warning, ToRelativePath(current.Path)));
            }
        }
        return candidates;
    }

    private void ValidateCatalog(CancellationToken cancellationToken) {
        List<MailboxCandidate> current = ScanCatalog(cancellationToken, false, out HashSet<string> folders);
        if (!folders.SetEquals(_catalogFolderPaths) || current.Count != _files.Count ||
            !new HashSet<string>(current.Select(file => file.RelativePath), StringComparer.Ordinal)
                .SetEquals(_files.Select(file => file.RelativePath))) {
            throw new InvalidDataException("The mailbox-directory catalog changed after it was indexed; reopen the session.");
        }
    }

    private void EnsureFolder(string path, IReadOnlyDictionary<string, int> counts) {
        if (_foldersById.ContainsKey(GetFolderId(path))) return;
        string? parentPath = GetParentPath(path);
        if (parentPath == null && path != "." && _catalogFolderPaths.Contains(".")) parentPath = ".";
        if (parentPath != null) EnsureFolder(parentPath, counts);
        if (_folders.Count >= _options.MaxFolderCount) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxFolderCount),
                _folders.Count + 1L, _options.MaxFolderCount);
        }
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
        if (IsAggregateMailbox(file)) return true;
        string extension = file.Extension;
        if (extension.Equals(".emlx", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".eml", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".mime", StringComparison.OrdinalIgnoreCase)) return true;
        string? parent = file.Directory?.Name;
        return string.Equals(parent, "cur", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(parent, "new", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAggregateMailbox(FileInfo file) =>
        string.Equals(file.Name, "mbox", StringComparison.OrdinalIgnoreCase) &&
        file.Directory?.Name.EndsWith(".mbox", StringComparison.OrdinalIgnoreCase) == true;
}
