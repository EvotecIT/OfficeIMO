using OfficeIMO.Drawing.Internal;
using OfficeIMO.Email;
using System.Collections.ObjectModel;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStorePstMutationTransaction {
    /// <summary>
    /// Writes and verifies a staged Unicode PST, optionally commits a byte-for-byte backup, and atomically
    /// replaces the original. Any failure before the final commit leaves the original unchanged.
    /// </summary>
    public EmailStorePstMutationReport Commit(CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        EnsureItemsIndexed(cancellationToken);
        if (!HasChanges()) {
            throw new InvalidOperationException("No effective PST mutations have been staged.");
        }

        string stagingPath = OfficeFileCommit.CreateStagingPath(_sourcePath);
        string? backupStagingPath = null;
        try {
            var folderMap = new Dictionary<string, string>(StringComparer.Ordinal);
            var folderParentMap = new Dictionary<string, string?>(StringComparer.Ordinal);
            var itemMap = new Dictionary<string, string>(StringComparer.Ordinal);
            var verificationMappings = new List<VerificationMapping>();
            EmailStorePstWriteReport writeReport;
            var writerOptions = new EmailStorePstWriterOptions(
                _source!.DisplayName,
                overwriteExisting: false,
                // The transaction owns the final fidelity gate so it can combine late source-read
                // diagnostics with writer diagnostics before deciding whether replacement is safe.
                failOnDataLoss: false,
                maxFolderCount: _options.MaxFolderCount,
                maxItemCount: _options.MaxItemCount,
                maxNestedMessageDepth: _options.MaxNestedMessageDepth,
                retainCheckpointOnDispose: false);
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(stagingPath, writerOptions)) {
                BuildFolderMap(writer, folderMap, folderParentMap);
                FailOnFidelityDiagnostics();
                var readOptions = new EmailStoreItemReadOptions(
                    EmailStoreItemReadParts.All, preferStreamingAttachmentContent: true);
                foreach (ItemState item in _items!.Values.OrderBy(value => value.Id, StringComparer.Ordinal)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (item.Deleted) continue;
                    if (!folderMap.TryGetValue(item.FolderId, out string? destinationFolder)) {
                        throw new InvalidDataException("An active item references an unmapped folder.");
                    }
                    EmailDocument document = item.Document ??
                        _source.ReadItem(item.Source!, readOptions, cancellationToken).Document;
                    string destinationItem = writer.AddItem(destinationFolder, document,
                        item.IsAssociated, cancellationToken);
                    itemMap.Add(item.Id, destinationItem);
                    verificationMappings.Add(new VerificationMapping(item,
                        destinationFolder, destinationItem));
                    if (item.Source?.IsOrphaned == true) {
                        _diagnostics.Add(new EmailStoreDiagnostic(
                            "EMAIL_STORE_PST_MUTATE_ORPHAN_RECOVERED",
                            "An item absent from its source contents table was recovered from the source index and retained.",
                            EmailStoreDiagnosticSeverity.Information, item.Id));
                    }
                }
                writeReport = writer.Complete(cancellationToken);
                _diagnostics.AddRange(writeReport.Diagnostics);
            }

            FailOnFidelityDiagnostics();
            EmailStorePstMutationVerificationReport? verification = _options.VerifyAfterWrite
                ? VerifyStagedPst(stagingPath, folderMap, folderParentMap,
                    verificationMappings, cancellationToken)
                : null;
            if (verification?.IsSuccessful == false) {
                throw new InvalidDataException(
                    "The staged PST did not match the intended folder and item semantics; the original was not changed.");
            }
            FailOnFidelityDiagnostics();
            cancellationToken.ThrowIfCancellationRequested();

            string? committedBackupPath = null;
            if (_options.BackupPath != null) {
                backupStagingPath = OfficeFileCommit.CreateStagingPath(_options.BackupPath);
                OfficeFileCommit.EnsureTargetDirectory(_options.BackupPath);
                File.Copy(_sourcePath, backupStagingPath, overwrite: false);
                cancellationToken.ThrowIfCancellationRequested();
                OfficeFileCommit.CommitTemporaryFile(backupStagingPath, _options.BackupPath,
                    _options.OverwriteBackup
                        ? OfficeFileCommit.ConflictPolicy.Replace
                        : OfficeFileCommit.ConflictPolicy.FailIfExists);
                backupStagingPath = null;
                committedBackupPath = _options.BackupPath;
            }

            cancellationToken.ThrowIfCancellationRequested();
            using (var commitGuard = new FileStream(_sourcePath, FileMode.Open, FileAccess.Read,
                FileShare.Read | FileShare.Delete, 1, FileOptions.RandomAccess)) {
                _source.Dispose();
                _source = null;
                EnsureSourceUnchanged();
                // FileShare.Delete is required for the atomic replacement itself. The adjacent
                // OfficeIMO lock owns pathname coordination; uncooperative replacers remain an
                // explicitly documented filesystem boundary rather than being silently assumed safe.
                OfficeFileCommit.CommitTemporaryFile(stagingPath, _sourcePath,
                    OfficeFileCommit.ConflictPolicy.Replace);
            }
            stagingPath = string.Empty;
            _committed = true;
            writeReport = new EmailStorePstWriteReport(_sourcePath, writeReport.FolderCount,
                writeReport.ItemCount, new FileInfo(_sourcePath).Length,
                writeReport.Diagnostics, writeReport.DiagnosticsTruncated);

            int createdFolders = _folders.Values.Count(folder => folder.IsCreated && !folder.Deleted);
            int renamedFolders = _folders.Values.Count(folder => !folder.IsCreated && !folder.Deleted &&
                !string.Equals(folder.Name, folder.OriginalName, StringComparison.Ordinal));
            int movedFolders = _folders.Values.Count(folder => !folder.IsCreated && !folder.Deleted &&
                !string.Equals(folder.ParentId, folder.OriginalParentId, StringComparison.Ordinal));
            int deletedFolders = _folders.Values.Count(folder => !folder.IsCreated && folder.Deleted);
            int addedItems = _items.Values.Count(item => item.IsCreated && !item.Deleted);
            int replacedItems = _items.Values.Count(item => !item.IsCreated && !item.Deleted && item.Replaced);
            int movedItems = _items.Values.Count(item => !item.IsCreated && !item.Deleted &&
                (!string.Equals(item.FolderId, item.OriginalFolderId, StringComparison.Ordinal) ||
                 item.IsAssociated != item.OriginalIsAssociated));
            int deletedItems = _items.Values.Count(item => !item.IsCreated && item.Deleted);
            var report = new EmailStorePstMutationReport(_sourcePath, committedBackupPath,
                writeReport, verification, createdFolders, renamedFolders, movedFolders,
                deletedFolders, addedItems, replacedItems, movedItems, deletedItems,
                new ReadOnlyDictionary<string, string>(folderMap),
                new ReadOnlyDictionary<string, string>(itemMap),
                _diagnostics.AsReadOnly());
            _transactionLock?.Dispose();
            _transactionLock = null;
            return report;
        } catch {
            Dispose();
            throw;
        } finally {
            OfficeFileCommit.DeleteIfExists(stagingPath);
            OfficeFileCommit.DeleteIfExists(backupStagingPath);
        }
    }

    private void BuildFolderMap(EmailStorePstWriter writer,
        IDictionary<string, string> folderMap,
        IDictionary<string, string?> folderParentMap) {
        foreach (FolderState folder in _folders.Values.Where(folder => !folder.Deleted)) {
            if (folder.IsMappedSystemFolder) {
                folderMap[folder.Id] = writer.SpamSearchFolderId;
                folderParentMap[folder.Id] = writer.MessageStoreRootFolderId;
                writer.ConfigureFolderMetadata(writer.SpamSearchFolderId,
                    folder.Name, folder.ContainerClass);
                continue;
            }
            switch (folder.SpecialFolderKind) {
                case EmailStoreSpecialFolderKind.Root:
                    folderMap[folder.Id] = writer.MessageStoreRootFolderId;
                    folderParentMap[folder.Id] = null;
                    writer.ConfigureFolderMetadata(writer.MessageStoreRootFolderId,
                        folder.Name, folder.ContainerClass);
                    break;
                case EmailStoreSpecialFolderKind.IpmSubtree:
                    folderMap[folder.Id] = writer.RootFolderId;
                    folderParentMap[folder.Id] = writer.MessageStoreRootFolderId;
                    writer.ConfigureFolderMetadata(writer.RootFolderId,
                        folder.Name, folder.ContainerClass);
                    break;
                case EmailStoreSpecialFolderKind.DeletedItems:
                    folderMap[folder.Id] = writer.DeletedItemsFolderId;
                    folderParentMap[folder.Id] = writer.RootFolderId;
                    writer.ConfigureFolderMetadata(writer.DeletedItemsFolderId,
                        folder.Name, folder.ContainerClass);
                    break;
                case EmailStoreSpecialFolderKind.SearchRoot:
                    folderMap[folder.Id] = writer.SearchRootFolderId;
                    folderParentMap[folder.Id] = writer.MessageStoreRootFolderId;
                    writer.ConfigureFolderMetadata(writer.SearchRootFolderId,
                        folder.Name, folder.ContainerClass);
                    break;
            }
        }

        var pending = _folders.Values.Where(folder => !folder.Deleted &&
            !folderMap.ContainsKey(folder.Id)).ToList();
        bool progress;
        do {
            progress = false;
            for (int index = pending.Count - 1; index >= 0; index--) {
                FolderState folder = pending[index];
                string parent;
                if (folder.ParentId == null) parent = writer.RootFolderId;
                else if (!folderMap.TryGetValue(folder.ParentId, out parent!)) continue;
                EmailStoreSpecialFolderKind role = PstStoreWriterCore.CanAssignUserSpecialFolder(
                    folder.SpecialFolderKind)
                        ? folder.SpecialFolderKind
                        : EmailStoreSpecialFolderKind.Unknown;
                if (folder.ClassificationSource == EmailStoreFolderClassificationSource.SourceIdentifier &&
                    folder.SpecialFolderKind != EmailStoreSpecialFolderKind.Unknown &&
                    !PstStoreWriterCore.SupportsSpecialFolderKind(folder.SpecialFolderKind)) {
                    _diagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_PST_MUTATE_SPECIAL_FOLDER_UNSUPPORTED",
                        "A source-identified special-folder role cannot be authored by the managed PST writer.",
                        EmailStoreDiagnosticSeverity.Warning, folder.Id));
                }
                folderMap[folder.Id] = role == EmailStoreSpecialFolderKind.Unknown
                    ? writer.AddFolder(folder.Name, parent, folder.ContainerClass)
                    : writer.AddFolder(folder.Name, role, parent, folder.ContainerClass);
                folderParentMap[folder.Id] = parent;
                if (folder.IsSearchFolder) {
                    _diagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_PST_MUTATE_SEARCH_FOLDER_STATIC",
                        "A source search folder is retained as a static folder because its dynamic search definition cannot be regenerated.",
                        EmailStoreDiagnosticSeverity.Warning, folder.Id));
                }
                pending.RemoveAt(index);
                progress = true;
            }
        } while (progress && pending.Count > 0);

        foreach (FolderState folder in pending) {
            EmailStoreSpecialFolderKind role = PstStoreWriterCore.CanAssignUserSpecialFolder(
                folder.SpecialFolderKind)
                    ? folder.SpecialFolderKind
                    : EmailStoreSpecialFolderKind.Unknown;
            folderMap[folder.Id] = role == EmailStoreSpecialFolderKind.Unknown
                ? writer.AddFolder(folder.Name, writer.RootFolderId, folder.ContainerClass)
                : writer.AddFolder(folder.Name, role, writer.RootFolderId, folder.ContainerClass);
            folderParentMap[folder.Id] = writer.RootFolderId;
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_MUTATE_FOLDER_PARENT_RECOVERED",
                "A folder with an unavailable or cyclic parent was attached to the destination root.",
                EmailStoreDiagnosticSeverity.Warning, folder.Id));
        }
    }

    private EmailStorePstMutationVerificationReport VerifyStagedPst(string stagingPath,
        IReadOnlyDictionary<string, string> folderMap,
        IReadOnlyDictionary<string, string?> folderParentMap,
        IReadOnlyList<VerificationMapping> mappings,
        CancellationToken cancellationToken) {
        int matchedFolders = 0;
        int mismatchedFolders = 0;
        int failedFolders = 0;
        int matchedItems = 0;
        int mismatchedItems = 0;
        int failedItems = 0;
        bool issuesTruncated = false;
        var issues = new List<EmailStorePstMutationVerificationIssue>();
        var readerOptions = new EmailStoreReaderOptions(
            maxFolderCount: _options.MaxFolderCount,
            maxItemCount: _options.MaxItemCount,
            retainAttachmentContent: false,
            includeAssociatedItems: true,
            includeOrphanedItems: true,
            maxNestedMessageDepth: _options.MaxNestedMessageDepth);
        using (EmailStoreSession destination = EmailStoreSession.Open(
            stagingPath, readerOptions, cancellationToken)) {
            var destinationFolders = destination.Folders.ToDictionary(
                folder => folder.Id, StringComparer.Ordinal);
            foreach (FolderState folder in _folders.Values.Where(folder => !folder.Deleted)) {
                cancellationToken.ThrowIfCancellationRequested();
                string destinationId = folderMap[folder.Id];
                if (!destinationFolders.TryGetValue(destinationId, out EmailStoreFolderInfo? actual)) {
                    failedFolders++;
                    AddIssue(issues, ref issuesTruncated,
                        new EmailStorePstMutationVerificationIssue(
                            EmailStorePstMutationVerificationEntity.Folder,
                            folder.Id, destinationId,
                            "EMAIL_STORE_PST_MUTATE_VERIFY_FOLDER_MISSING",
                            Array.Empty<EmailSemanticDifference>()));
                    continue;
                }
                string? expectedParent = folderParentMap[folder.Id];
                bool metadataMatches =
                    string.Equals(actual.Name, folder.Name, StringComparison.Ordinal) &&
                    string.Equals(actual.ParentId, expectedParent, StringComparison.Ordinal) &&
                    string.Equals(actual.ContainerClass, folder.ContainerClass,
                        StringComparison.OrdinalIgnoreCase);
                bool verifySpecialRole = PstStoreWriterCore.SupportsSpecialFolderKind(
                    folder.SpecialFolderKind);
                bool specialRoleMatches = !verifySpecialRole ||
                    actual.SpecialFolderKind == folder.SpecialFolderKind &&
                    actual.ClassificationSource == EmailStoreFolderClassificationSource.SourceIdentifier;
                bool matches = metadataMatches && specialRoleMatches;
                if (matches) matchedFolders++;
                else {
                    mismatchedFolders++;
                    AddIssue(issues, ref issuesTruncated,
                        new EmailStorePstMutationVerificationIssue(
                            EmailStorePstMutationVerificationEntity.Folder,
                            folder.Id, destinationId,
                            "EMAIL_STORE_PST_MUTATE_VERIFY_FOLDER_MISMATCH",
                            Array.Empty<EmailSemanticDifference>()));
                }
            }

            var readOptions = new EmailStoreItemReadOptions(
                EmailStoreItemReadParts.All, preferStreamingAttachmentContent: true);
            EmailSemanticComparisonOptions comparisonOptions = _options.VerificationOptions ??
                new EmailSemanticComparisonOptions(
                    maxEmbeddedMessageDepth: _options.MaxNestedMessageDepth);
            foreach (VerificationMapping mapping in mappings) {
                cancellationToken.ThrowIfCancellationRequested();
                try {
                    EmailDocument expected = mapping.Item.Document ??
                        _source!.ReadItem(mapping.Item.Source!, readOptions, cancellationToken).Document;
                    var destinationReference = new EmailStoreItemReference(
                        mapping.DestinationItemId, mapping.DestinationFolderId,
                        mapping.Item.IsAssociated, isOrphaned: false);
                    EmailStoreItem actual = destination.ReadItem(
                        destinationReference, readOptions, cancellationToken);
                    EmailSemanticComparisonReport comparison = EmailSemanticComparer.Compare(
                        expected, actual.Document, comparisonOptions, cancellationToken);
                    if (comparison.IsMatch) matchedItems++;
                    else {
                        mismatchedItems++;
                        AddIssue(issues, ref issuesTruncated,
                            new EmailStorePstMutationVerificationIssue(
                                EmailStorePstMutationVerificationEntity.Item,
                                mapping.Item.Id, mapping.DestinationItemId,
                                "EMAIL_STORE_PST_MUTATE_VERIFY_ITEM_MISMATCH",
                                comparison.Differences));
                    }
                } catch (Exception exception) when (
                    exception is InvalidDataException || exception is IOException ||
                    exception is NotSupportedException || exception is KeyNotFoundException ||
                    exception is EmailStoreLimitExceededException ||
                    exception is EmailLimitExceededException) {
                    failedItems++;
                    AddIssue(issues, ref issuesTruncated,
                        new EmailStorePstMutationVerificationIssue(
                            EmailStorePstMutationVerificationEntity.Item,
                            mapping.Item.Id, mapping.DestinationItemId,
                            "EMAIL_STORE_PST_MUTATE_VERIFY_ITEM_FAILED",
                            Array.Empty<EmailSemanticDifference>()));
                }
            }
            _diagnostics.AddRange(destination.Diagnostics);
        }

        if (mismatchedFolders > 0 || failedFolders > 0) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_MUTATE_VERIFY_FOLDER_FAILED",
                "One or more staged folders did not match the intended hierarchy and metadata.",
                EmailStoreDiagnosticSeverity.Error, stagingPath));
        }
        if (mismatchedItems > 0 || failedItems > 0) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_MUTATE_VERIFY_ITEM_FAILED",
                "One or more staged items could not be reopened with the intended semantic projection.",
                EmailStoreDiagnosticSeverity.Error, stagingPath));
        }
        return new EmailStorePstMutationVerificationReport(
            _folders.Values.Count(folder => !folder.Deleted),
            matchedFolders, mismatchedFolders, failedFolders,
            mappings.Count, matchedItems, mismatchedItems, failedItems,
            issues.AsReadOnly(), issuesTruncated);
    }

    private void AddIssue(ICollection<EmailStorePstMutationVerificationIssue> issues,
        ref bool truncated, EmailStorePstMutationVerificationIssue issue) {
        if (issues.Count < _options.MaxVerificationIssues) issues.Add(issue);
        else truncated = true;
    }

    private void FailOnFidelityDiagnostics() {
        CaptureSourceDiagnostics();
        if (!_options.FailOnDataLoss) return;
        EmailStoreDiagnostic[] fidelity = _diagnostics.Where(diagnostic =>
            diagnostic.Severity != EmailStoreDiagnosticSeverity.Information).ToArray();
        if (fidelity.Length > 0) {
            throw new InvalidOperationException(
                string.Concat("The staged PST emitted a fidelity diagnostic and FailOnDataLoss is enabled; ",
                    "the original was not changed. Codes: ",
                    string.Join(", ", fidelity.Select(diagnostic => diagnostic.Code).Distinct(StringComparer.Ordinal))));
        }
    }

    private void CaptureSourceDiagnostics() {
        if (_source == null) return;
        foreach (EmailStoreDiagnostic diagnostic in _source.Diagnostics) {
            if (_diagnostics.Any(existing =>
                existing.Severity == diagnostic.Severity &&
                string.Equals(existing.Code, diagnostic.Code, StringComparison.Ordinal) &&
                string.Equals(existing.Message, diagnostic.Message, StringComparison.Ordinal) &&
                string.Equals(existing.Location, diagnostic.Location, StringComparison.Ordinal))) {
                continue;
            }
            _diagnostics.Add(diagnostic);
        }
    }

    private void EnsureSourceUnchanged() {
        var source = new FileInfo(_sourcePath);
        if (!source.Exists || source.Length != _sourceLength ||
            source.LastWriteTimeUtc != _sourceLastWriteTimeUtc) {
            throw new IOException(
                "The source PST changed while the mutation transaction was open; the staged rewrite was not committed.");
        }
    }

    private sealed class VerificationMapping {
        internal VerificationMapping(ItemState item, string destinationFolderId,
            string destinationItemId) {
            Item = item;
            DestinationFolderId = destinationFolderId;
            DestinationItemId = destinationItemId;
        }

        internal ItemState Item { get; }
        internal string DestinationFolderId { get; }
        internal string DestinationItemId { get; }
    }
}
