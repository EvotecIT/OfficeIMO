using OfficeIMO.Drawing.Internal;
using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStorePstMutationTransaction {
    /// <summary>
    /// Writes and verifies a staged Unicode PST, asynchronously flushes the staged artifact and optional byte-for-byte
    /// backup, then atomically replaces the original. Ordered PST serialization and semantic comparison remain
    /// synchronous stateful phases and are not hidden in a worker-thread wrapper.
    /// </summary>
    public async Task<EmailStorePstMutationReport> CommitAsync(
        CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        EnsureItemsIndexed(cancellationToken);
        if (!HasChanges()) {
            throw new InvalidOperationException("No effective PST mutations have been staged.");
        }
        EmailStorePstMutationPlan plan = DryRun(cancellationToken);

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
                failOnDataLoss: false,
                maxFolderCount: _options.MaxFolderCount,
                maxItemCount: _options.MaxItemCount,
                maxNestedMessageDepth: _options.MaxNestedMessageDepth,
                retainCheckpointOnDispose: false);
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(stagingPath, writerOptions)) {
                if (!_folders.Values.Any(folder => !folder.Deleted && folder.IsMappedSystemFolder)) {
                    writer.SuppressWriterOwnedSpamSearchFolder();
                }
                BuildFolderMap(writer, folderMap, folderParentMap, cancellationToken);
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

            await FlushFileAsync(stagingPath, cancellationToken).ConfigureAwait(false);
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
            cancellationToken.ThrowIfCancellationRequested();
            using (var commitGuard = new FileStream(_sourcePath, FileMode.Open, FileAccess.Read,
                FileShare.Read | FileShare.Delete, 1, FileOptions.RandomAccess)) {
                _source.Dispose();
                _source = null;
                EnsureSourceUnchanged();
                if (_options.BackupPath != null) {
                    backupStagingPath = OfficeFileCommit.CreateStagingPath(_options.BackupPath);
                    OfficeFileCommit.EnsureTargetDirectory(_options.BackupPath);
                    await CopyFileAsync(_sourcePath, backupStagingPath, cancellationToken).ConfigureAwait(false);
                    cancellationToken.ThrowIfCancellationRequested();
                    EnsureSourceUnchanged();
                    OfficeFileCommit.CommitTemporaryFileAtomically(backupStagingPath, _options.BackupPath,
                        _options.OverwriteBackup
                            ? OfficeFileCommit.ConflictPolicy.Replace
                            : OfficeFileCommit.ConflictPolicy.FailIfExists);
                    backupStagingPath = null;
                    committedBackupPath = _options.BackupPath;
                }
                OfficeFileCommit.CommitTemporaryFileAtomically(stagingPath, _sourcePath,
                    OfficeFileCommit.ConflictPolicy.Replace);
            }
            stagingPath = string.Empty;
            _committed = true;
            writeReport = new EmailStorePstWriteReport(_sourcePath, writeReport.FolderCount,
                writeReport.ItemCount, new FileInfo(_sourcePath).Length,
                writeReport.Diagnostics, writeReport.DiagnosticsTruncated);
            return CreateCommittedReport(committedBackupPath, plan, writeReport, verification,
                folderMap, itemMap);
        } catch {
            Dispose();
            throw;
        } finally {
            OfficeFileCommit.DeleteIfExists(stagingPath);
            OfficeFileCommit.DeleteIfExists(backupStagingPath);
        }
    }

    private static async Task FlushFileAsync(string path, CancellationToken cancellationToken) {
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.Read,
                   128 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan)) {
            await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
    }

    private static async Task CopyFileAsync(string sourcePath, string destinationPath,
        CancellationToken cancellationToken) {
        using (var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read,
                   FileShare.Read | FileShare.Delete, 128 * 1024,
                   FileOptions.Asynchronous | FileOptions.SequentialScan))
        using (var destination = new FileStream(destinationPath, FileMode.CreateNew, FileAccess.Write,
                   FileShare.None, 128 * 1024,
                   FileOptions.Asynchronous | FileOptions.SequentialScan)) {
            await source.CopyToAsync(destination, 128 * 1024, cancellationToken).ConfigureAwait(false);
            await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
            if (destination.Length != source.Length) {
                throw new IOException("The staged PST backup length does not match the source.");
            }
        }
    }
}
