using OfficeIMO.Core.Internal;

namespace OfficeIMO.Email.Store;

internal sealed partial class MailboxDirectoryStoreSessionBackend {
    private EmailStoreSourceGuard? _dbxCatalogGuard;
    private bool _dbxFileSetInvalid;
    private bool _dbxBulkReadActive;
    private readonly List<string> _dbxMissingFiles = new List<string>();
    private readonly List<DbxStoreSessionBackend> _dbxMailboxes = new List<DbxStoreSessionBackend>();

    internal IDisposable? BeginDbxBulkRead(CancellationToken cancellationToken) {
        if (_dbxCatalogPath == null) return null;
        if (_dbxBulkReadActive) throw new InvalidOperationException("A DBX catalog bulk read is already active.");
        ValidateDbxCatalog(cancellationToken);
        var scopes = new List<IDisposable>(_dbxMailboxes.Count);
        try {
            foreach (DbxStoreSessionBackend mailbox in _dbxMailboxes) {
                scopes.Add(mailbox.BeginBulkRead(cancellationToken));
            }
            var scope = new DbxBulkReadScope(this, scopes, cancellationToken);
            _dbxBulkReadActive = true;
            return scope;
        } catch {
            foreach (IDisposable scope in scopes) {
                // Reset every admitted child while preserving the original admission failure.
                try { scope.Dispose(); } catch { }
            }
            throw;
        }
    }

    // An Outlook Express catalog names sibling DBX files. The existing directory owner
    // enforces regular-file access, aggregate bounds, fingerprints and export containment.
    private void IndexDbxCatalog(string catalogPath, CancellationToken cancellationToken) {
        using FileStream source = OpenDbxFile(catalogPath);
        string catalogName = Path.GetFileName(catalogPath);
        var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { catalogName };
        RegisterDbxFile(catalogPath, catalogName, string.Empty, source.Length);
        _dbxCatalogGuard = new EmailStoreSourceGuard(source, source.Length, _options.MaxInputBytes,
            _resources.Dispose, cancellationToken);

        var database = new DbxDatabase(source, _options);
        var folders = new Dictionary<uint, DbxFolder>();
        foreach (uint address in database.ReadIndex(folderDatabase: true, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            DbxIndexedRecord record = database.ReadIndexedRecord(address, cancellationToken);
            uint? sourceType = record.ReadUInt32(0x0a, cancellationToken);
            if (sourceType != 3) {
                AddDiagnostic(new EmailStoreDiagnostic("EMAIL_STORE_DBX_NONLOCAL_FOLDER_SKIPPED",
                    "The Outlook Express catalog contains a non-local folder outside the local mail profile.",
                    EmailStoreDiagnosticSeverity.Information));
                continue;
            }

            uint id = record.ReadUInt32(0x00, cancellationToken) ?? 0;
            uint parent = record.ReadUInt32(0x01, cancellationToken) ?? 0;
            if (id == 0 || folders.ContainsKey(id))
                throw DbxDatabase.Invalid("The active local folders have an invalid or repeated identifier.");
            string name = record.ReadString(0x02, cancellationToken) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(name))
                throw DbxDatabase.Invalid("An active local folder has no display name.");
            string? fileName = record.ReadString(0x03, cancellationToken);
            if (fileName?.Length == 0) fileName = null;
            if (fileName != null && (!IsSafeDbxFileName(fileName) || !files.Add(fileName)))
                throw DbxDatabase.Invalid("An active local folder refers to an unsafe or repeated mailbox file name.");
            uint? declaredCount = record.ReadUInt32(0x07, cancellationToken);
            EmailStoreSpecialFolderKind role = GetDbxFolderRole(record.ReadUInt32(0x09, cancellationToken));
            folders.Add(id, new DbxFolder(id, parent, name, fileName, declaredCount, role));
        }

        ValidateDbxParents(folders, cancellationToken);
        foreach (DbxFolder folder in folders.Values.OrderBy(value => value.Id)) {
            cancellationToken.ThrowIfCancellationRequested();
            string id = DbxFolderId(folder.Id);
            int? count = IndexDbxFolder(folder, id, cancellationToken);
            string? parent = folder.Parent == 0 || folder.Parent == uint.MaxValue
                ? null : DbxFolderId(folder.Parent);
            EmailStoreFolderClassificationSource classification = folder.Role == EmailStoreSpecialFolderKind.Unknown
                ? EmailStoreFolderClassificationSource.None : EmailStoreFolderClassificationSource.SourceIdentifier;
            var info = new EmailStoreFolderInfo(id, parent, folder.Name, count, 0, folder.Role, classification);
            _folders.Add(info);
            _foldersById.Add(id, info);
        }
        _dbxCatalogGuard.Validate(source, cancellationToken);
    }

    private int? IndexDbxFolder(DbxFolder folder, string id, CancellationToken cancellationToken) {
        if (folder.FileName == null) {
            if (folder.DeclaredCount.GetValueOrDefault() == 0) return 0;
            ReportMissingDbxMailbox(id);
            return null;
        }

        string path = Path.Combine(_root, folder.FileName);
        if (!File.Exists(path)) {
            _dbxMissingFiles.Add(path);
            ReportMissingDbxMailbox(id);
            return null;
        }
        // Admit the aggregate size and file count using the same handle that will be indexed.
        using FileStream input = OpenDbxFile(path);
        RegisterDbxFile(path, folder.FileName, id, input.Length);
        var backend = new DbxStoreSessionBackend(input, () => OpenDbxFile(path), folder.FileName,
            _options, cancellationToken, _resources, _items.Count);
        _dbxMailboxes.Add(backend);
        int count = 0;
        foreach (EmailStoreItemReference inner in backend.EnumerateItems(new EmailStoreEnumerationOptions(),
                     cancellationToken)) {
            string itemId = id + ":" + inner.Id;
            AddItem(new EmailStoreItemReference(itemId, id, false, false), cancellationToken);
            _aggregateItems.Add(itemId, new AggregateItem(backend, inner, id, folder.FileName));
            count++;
        }

        if (folder.DeclaredCount.HasValue && folder.DeclaredCount.Value != count) {
            AddDiagnostic(new EmailStoreDiagnostic("EMAIL_STORE_DBX_FOLDER_COUNT_MISMATCH",
                "The folder catalog count differs from the selected mailbox's active message index.",
                EmailStoreDiagnosticSeverity.Warning, id));
        }
        return count;
    }

    private void RegisterDbxFile(string path, string name, string folderId, long length) {
        if (_files.Count >= _options.MaxDirectoryFileCount)
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDirectoryFileCount),
                _files.Count + 1L, _options.MaxDirectoryFileCount);
        _sourceLength = AddBounded(_sourceLength, length);
        _files.Add(new MailboxFile("dbx:file:" + name, folderId, path, name, false, null, length));
    }

    private void ReportMissingDbxMailbox(string id) =>
        AddDiagnostic(new EmailStoreDiagnostic("EMAIL_STORE_DBX_MAILBOX_MISSING",
            "An active local folder refers to a mailbox whose local bytes are unavailable.",
            EmailStoreDiagnosticSeverity.Error, id, "index-dbx-catalog", null, null, null, null,
            EmailDiagnosticDisposition.Skipped, EmailDataLossRisk.Possible,
            "Provide the complete original Outlook Express store directory and reopen the catalog."));

    private void ValidateDbxCatalog(CancellationToken cancellationToken) {
        if (_dbxCatalogPath == null) return;
        using FileStream source = OpenDbxFile(_dbxCatalogPath);
        _dbxCatalogGuard!.Validate(source, cancellationToken, verifyFingerprint: !_dbxBulkReadActive);
        foreach (string missingPath in _dbxMissingFiles) {
            cancellationToken.ThrowIfCancellationRequested();
            if (File.Exists(missingPath))
                throw InvalidateDbxFileSet("The referenced mailbox set changed after the catalog was opened.");
        }
    }

    private FileStream OpenDbxFile(string path) {
        if (_dbxFileSetInvalid) throw DbxDatabase.Invalid("Reopen the catalog after a referenced file changed.");
        try {
            if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0)
                throw InvalidateDbxFileSet("A referenced DBX file is a symbolic link or reparse point.");
            return OpenRegularMailboxFile(path);
        } catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException) {
            throw InvalidateDbxFileSet("A referenced DBX file is no longer a regular readable file.", exception);
        }
    }

    private InvalidDataException InvalidateDbxFileSet(string message, Exception? exception = null) {
        _dbxFileSetInvalid = true;
        _resources.Dispose();
        return new InvalidDataException("Invalid Outlook Express DBX. " + message, exception);
    }

    private void ValidateDbxParents(Dictionary<uint, DbxFolder> folders, CancellationToken cancellationToken) {
        foreach (DbxFolder folder in folders.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            uint parent = folder.Parent;
            int depth = 0;
            var seen = new HashSet<uint> { folder.Id };
            while (parent != 0 && parent != uint.MaxValue) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!seen.Add(parent))
                    throw DbxDatabase.Invalid("The active local folder hierarchy contains a cycle.");
                if (++depth > _options.MaxDirectoryDepth)
                    throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDirectoryDepth),
                        depth, _options.MaxDirectoryDepth);
                if (!folders.TryGetValue(parent, out DbxFolder? ancestor))
                    throw DbxDatabase.Invalid("An active local folder refers to a missing local parent.");
                parent = ancestor.Parent;
            }
        }
    }

    private static bool IsSafeDbxFileName(string name) =>
        name.Length > 4 && name.EndsWith(".dbx", StringComparison.OrdinalIgnoreCase) &&
        string.Equals(OfficePortableFileName.SanitizeBaseName(name, maximumLength: 255), name,
            StringComparison.Ordinal);

    private static string DbxFolderId(uint id) => "dbx:folder:" + id.ToString("x8", CultureInfo.InvariantCulture);

    private static EmailStoreSpecialFolderKind GetDbxFolderRole(uint? role) => role switch {
        1 => EmailStoreSpecialFolderKind.Inbox,
        2 => EmailStoreSpecialFolderKind.Outbox,
        3 => EmailStoreSpecialFolderKind.SentItems,
        4 => EmailStoreSpecialFolderKind.DeletedItems,
        5 => EmailStoreSpecialFolderKind.Drafts,
        _ => EmailStoreSpecialFolderKind.Unknown
    };

    private sealed class DbxBulkReadScope : IDisposable {
        private readonly MailboxDirectoryStoreSessionBackend _owner;
        private readonly IReadOnlyList<IDisposable> _scopes;
        private readonly CancellationToken _cancellationToken;
        private bool _disposed;

        internal DbxBulkReadScope(MailboxDirectoryStoreSessionBackend owner,
            IReadOnlyList<IDisposable> scopes, CancellationToken cancellationToken) {
            _owner = owner;
            _scopes = scopes;
            _cancellationToken = cancellationToken;
        }

        public void Dispose() {
            if (_disposed) return;
            _disposed = true;
            _owner._dbxBulkReadActive = false;
            Exception? failure = null;
            foreach (IDisposable scope in _scopes) {
                try { scope.Dispose(); }
                catch (Exception exception) { failure ??= exception; }
            }
            try { _owner.ValidateDbxCatalog(_cancellationToken); }
            catch (Exception exception) { failure ??= exception; }
            if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }

    private sealed class DbxFolder {
        internal DbxFolder(uint id, uint parent, string name, string? fileName, uint? declaredCount,
            EmailStoreSpecialFolderKind role) {
            Id = id;
            Parent = parent;
            Name = name;
            FileName = fileName;
            DeclaredCount = declaredCount;
            Role = role;
        }

        internal uint Id { get; }
        internal uint Parent { get; }
        internal string Name { get; }
        internal string? FileName { get; }
        internal uint? DeclaredCount { get; }
        internal EmailStoreSpecialFolderKind Role { get; }
    }
}
