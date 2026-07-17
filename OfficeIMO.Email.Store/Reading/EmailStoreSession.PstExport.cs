using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Streams this open store into a newly created Unicode PST. The source is read-only and is never mutated.
    /// </summary>
    public EmailStorePstConversionReport ExportToPst(string destinationPath,
        EmailStorePstConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(destinationPath)) {
            throw new ArgumentException("A destination path is required.", nameof(destinationPath));
        }
        ThrowIfDisposed();
        var effective = options ?? new EmailStorePstConversionOptions();
        string destination = Path.GetFullPath(destinationPath);
        ThrowIfStoreSourceDestination(destination, "PST export");
        ValidateVerificationManifestPath(destination, effective);
        if (File.Exists(destination) && !effective.OverwriteExisting) {
            throw new IOException("The destination PST already exists and overwriteExisting is false.");
        }

        var diagnostics = new List<EmailStoreDiagnostic>();
        string? stagingPath = effective.VerifyAfterWrite
            ? OfficeFileCommit.CreateStagingPath(destination)
            : null;
        string? manifestStagingPath = effective.VerificationManifestPath == null
            ? null
            : OfficeFileCommit.CreateTemporaryPath(
                Path.GetFullPath(effective.VerificationManifestPath));
        string writerDestination = stagingPath ?? destination;
        var writerOptions = new EmailStorePstWriterOptions(
            effective.DisplayName ?? DisplayName,
            stagingPath == null && effective.OverwriteExisting,
            effective.FailOnDataLoss,
            maxFolderCount: Math.Max(1, Folders.Count + 8),
            maxItemCount: effective.MaxItems,
            maxNestedMessageDepth: effective.MaxNestedMessageDepth);
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(writerDestination, writerOptions))
            using (PstConversionMappingJournal? mappings = effective.VerifyAfterWrite
                       ? new PstConversionMappingJournal(writerDestination)
                       : null) {
                Dictionary<string, string> folderMap = CreatePstFolderMap(writer, effective, diagnostics);
                int converted = 0;
                int skipped = 0;
                var enumeration = new EmailStoreEnumerationOptions(
                    includeAssociatedItems: effective.IncludeAssociatedItems,
                    includeOrphanedItems: effective.IncludeOrphanedItems,
                    maxItems: effective.MaxItems);
                var readOptions = new EmailStoreItemReadOptions(
                    EmailStoreItemReadParts.All, preferStreamingAttachmentContent: true);
                foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    EmailStoreFolderInfo? sourceFolder = Folders.FirstOrDefault(item => item.Id == reference.FolderId);
                    if (sourceFolder?.IsSearchFolder == true && !effective.IncludeSearchFolders) {
                        skipped++;
                        continue;
                    }
                    if (!folderMap.TryGetValue(reference.FolderId, out string? destinationFolder)) {
                        skipped++;
                        diagnostics.Add(new EmailStoreDiagnostic(
                            "EMAIL_STORE_PST_CONVERT_FOLDER_UNMAPPED",
                            "An item was skipped because its source folder could not be mapped.",
                            EmailStoreDiagnosticSeverity.Error, reference.Id));
                        if (!effective.ContinueOnItemError) {
                            throw new InvalidDataException("A source item folder could not be mapped.");
                        }
                        continue;
                    }
                    try {
                        EmailStoreItem item = ReadItem(reference, readOptions, cancellationToken);
                        string destinationItemId = writer.AddItem(destinationFolder, item.Document,
                            reference.IsAssociated, cancellationToken);
                        converted++;
                        if (effective.VerifyAfterWrite) {
                            mappings!.Add(converted, reference,
                                destinationFolder, destinationItemId);
                        }
                        if (reference.IsOrphaned) {
                            diagnostics.Add(new EmailStoreDiagnostic(
                                "EMAIL_STORE_PST_CONVERT_ORPHAN_RECOVERED",
                                "An item absent from its source contents table was recovered from the source index and copied.",
                                EmailStoreDiagnosticSeverity.Information, reference.Id));
                        }
                    } catch (Exception exception) when (effective.ContinueOnItemError &&
                        (exception is InvalidDataException || exception is NotSupportedException ||
                         exception is IOException || exception is EmailStoreLimitExceededException)) {
                        skipped++;
                        diagnostics.Add(new EmailStoreDiagnostic(
                            "EMAIL_STORE_PST_CONVERT_ITEM_SKIPPED",
                            string.Concat("A source item could not be copied: ", exception.Message),
                            EmailStoreDiagnosticSeverity.Error, reference.Id));
                    }
                }

                if (effective.FailOnDataLoss && diagnostics.Any(item =>
                    item.Severity != EmailStoreDiagnosticSeverity.Information)) {
                    throw new InvalidOperationException(
                        "Store conversion produced fidelity diagnostics and FailOnDataLoss is enabled.");
                }
                EmailStorePstWriteReport writeReport = writer.Complete(cancellationToken);
                diagnostics.AddRange(writeReport.Diagnostics);
                EmailStorePstVerificationReport? verification = effective.VerifyAfterWrite
                    ? VerifyPstConversion(writerDestination, mappings!, effective, diagnostics,
                        manifestStagingPath, cancellationToken)
                    : null;
                if (effective.FailOnDataLoss && verification?.IsSuccessful == false) {
                    throw new InvalidOperationException(
                        "PST conversion semantic verification reported data loss; the destination was not changed.");
                }
                if (stagingPath != null) {
                    OfficeFileCommit.CommitTemporaryFile(stagingPath, destination,
                        effective.OverwriteExisting
                            ? OfficeFileCommit.ConflictPolicy.Replace
                            : OfficeFileCommit.ConflictPolicy.FailIfExists);
                    stagingPath = null;
                    writeReport = new EmailStorePstWriteReport(destination, writeReport.FolderCount,
                        writeReport.ItemCount, new FileInfo(destination).Length,
                        writeReport.Diagnostics, writeReport.DiagnosticsTruncated);
                }
                if (manifestStagingPath != null) {
                    string manifestDestination = Path.GetFullPath(effective.VerificationManifestPath!);
                    OfficeFileCommit.CommitTemporaryFile(manifestStagingPath, manifestDestination,
                        effective.OverwriteExisting
                            ? OfficeFileCommit.ConflictPolicy.Replace
                            : OfficeFileCommit.ConflictPolicy.FailIfExists);
                    manifestStagingPath = null;
                    verification = verification!.WithManifestPath(manifestDestination);
                }
                return new EmailStorePstConversionReport(Format, writeReport,
                    Folders.Count, converted, skipped, verification, diagnostics.ToArray());
            }
        } finally {
            OfficeFileCommit.DeleteIfExists(stagingPath);
            OfficeFileCommit.DeleteIfExists(manifestStagingPath);
        }
    }

    private void ValidateVerificationManifestPath(string destination,
        EmailStorePstConversionOptions options) {
        if (options.VerificationManifestPath == null) return;
        string manifest = Path.GetFullPath(options.VerificationManifestPath);
        ThrowIfStoreSourceDestination(manifest, "PST verification manifest");
        if (string.Equals(manifest, destination, StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidOperationException(
                "The verification manifest and destination PST must use different paths.");
        }
        if (File.Exists(manifest) && !options.OverwriteExisting) {
            throw new IOException(
                "The verification manifest already exists and overwriteExisting is false.");
        }
    }

    private Dictionary<string, string> CreatePstFolderMap(EmailStorePstWriter writer,
        EmailStorePstConversionOptions options, IList<EmailStoreDiagnostic> diagnostics) {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (EmailStoreFolderInfo folder in Folders) {
            if (folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Root ||
                folder.SpecialFolderKind == EmailStoreSpecialFolderKind.IpmSubtree) {
                map[folder.Id] = writer.RootFolderId;
            } else if (folder.SpecialFolderKind == EmailStoreSpecialFolderKind.DeletedItems) {
                map[folder.Id] = writer.DeletedItemsFolderId;
            } else if (folder.SpecialFolderKind == EmailStoreSpecialFolderKind.SearchRoot) {
                map[folder.Id] = writer.SearchRootFolderId;
            }
        }

        var pending = Folders.Where(item => !map.ContainsKey(item.Id)).ToList();
        bool progress;
        do {
            progress = false;
            for (int index = pending.Count - 1; index >= 0; index--) {
                EmailStoreFolderInfo folder = pending[index];
                if (folder.IsSearchFolder && !options.IncludeSearchFolders) {
                    pending.RemoveAt(index);
                    progress = true;
                    continue;
                }
                string? parent;
                if (folder.ParentId == null) parent = writer.RootFolderId;
                else if (!map.TryGetValue(folder.ParentId, out parent)) continue;
                map[folder.Id] = writer.AddFolder(folder.Name, parent, folder.ContainerClass);
                if (folder.IsSearchFolder) {
                    diagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_PST_CONVERT_SEARCH_FOLDER_STATIC",
                        "A search folder was copied as a static folder; its dynamic search definition is not regenerated.",
                        EmailStoreDiagnosticSeverity.Warning, folder.Id));
                }
                pending.RemoveAt(index);
                progress = true;
            }
        } while (progress && pending.Count > 0);

        foreach (EmailStoreFolderInfo folder in pending) {
            map[folder.Id] = writer.AddFolder(folder.Name, writer.RootFolderId,
                folder.ContainerClass);
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_CONVERT_FOLDER_PARENT_RECOVERED",
                "A folder with an unavailable or cyclic parent was attached to the destination root.",
                EmailStoreDiagnosticSeverity.Warning, folder.Id));
        }
        return map;
    }
}
