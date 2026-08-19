using OfficeIMO.Core.Internal;

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
        ValidateMigrationCheckpointPath(destination, effective);
        bool resumeRequested = effective.CheckpointPath != null && File.Exists(effective.CheckpointPath);
        if (!resumeRequested && File.Exists(destination) && !effective.OverwriteExisting) {
            throw new IOException("The destination PST already exists and overwriteExisting is false.");
        }

        var sourceIdentity = new EmailStoreSourceIdentity(Format, SourceLength,
            GetCatalogFingerprint(cancellationToken), GetDurableSourceFingerprint(cancellationToken));
        string optionsFingerprint = effective.GetMigrationFingerprint(destination);
        var diagnostics = new List<EmailStoreDiagnostic>();
        string? stagingPath = null;
        string? manifestStagingPath = effective.VerificationManifestPath == null
            ? null
            : OfficeFileCommit.CreateTemporaryPath(
                Path.GetFullPath(effective.VerificationManifestPath));
        EmailStorePstWriter? writer = null;
        PstConversionMappingJournal? mappings = null;
        PstConversionCheckpointState state = null!;
        bool stateReady = false;
        bool writerCompleted = false;
        bool completionStarted = false;
        bool wasResumed = false;
        bool destinationMutationUncertain = false;
        try {
            if (resumeRequested) {
                writer = EmailStorePstWriter.Resume(effective.CheckpointPath!,
                    out byte[]? applicationState);
                state = PstConversionCheckpointState.Deserialize(applicationState ??
                    throw new InvalidDataException("The PST writer checkpoint has no migration state."));
                ValidateResumedMigration(state, sourceIdentity, optionsFingerprint, destination, effective);
                diagnostics.AddRange(state.Diagnostics);
                stagingPath = effective.VerifyAfterWrite ? state.WriterDestination : null;
                if (effective.VerifyAfterWrite) {
                    mappings = new PstConversionMappingJournal(state.MappingPath!, resume: true,
                        state.MappingLength, state.MappingCount);
                }
                wasResumed = true;
                stateReady = true;
            } else {
                stagingPath = effective.VerifyAfterWrite
                    ? OfficeFileCommit.CreateStagingPath(destination)
                    : null;
                string writerDestination = stagingPath ?? destination;
                var writerOptions = new EmailStorePstWriterOptions(
                    effective.DisplayName ?? DisplayName,
                    stagingPath == null && effective.OverwriteExisting,
                    effective.FailOnDataLoss,
                    maxFolderCount: Math.Max(1, Folders.Count + 8),
                    maxItemCount: effective.MaxItems,
                    maxNestedMessageDepth: effective.MaxNestedMessageDepth,
                    checkpointPath: effective.CheckpointPath,
                    checkpointIntervalItems: int.MaxValue,
                    retainCheckpointOnDispose: true);
                writer = EmailStorePstWriter.Create(writerDestination, writerOptions);
                Dictionary<string, string> createdFolders = CreatePstFolderMap(writer, effective, diagnostics);
                state = new PstConversionCheckpointState {
                    SourceFormat = sourceIdentity.Format,
                    SourceLength = sourceIdentity.Length,
                    CatalogFingerprint = sourceIdentity.CatalogFingerprint,
                    DurableFingerprint = sourceIdentity.DurableFingerprint,
                    OptionsFingerprint = optionsFingerprint,
                    WriterDestination = writerDestination
                };
                foreach (KeyValuePair<string, string> folder in createdFolders) {
                    state.FolderMap.Add(folder.Key, folder.Value);
                }
                if (effective.VerifyAfterWrite) {
                    mappings = effective.CheckpointPath == null
                        ? new PstConversionMappingJournal(writerDestination)
                        : new PstConversionMappingJournal(
                            string.Concat(effective.CheckpointPath, ".verify-map"),
                            resume: false, committedLength: 0, committedCount: 0);
                    state.MappingPath = mappings.Path;
                }
                if (effective.CheckpointPath != null) {
                    CommitMigrationCheckpoint(writer, mappings, state, diagnostics);
                }
                stateReady = true;
            }

            var enumeration = new EmailStoreEnumerationOptions(
                includeAssociatedItems: effective.IncludeAssociatedItems,
                includeOrphanedItems: effective.IncludeOrphanedItems,
                maxItems: effective.MaxItems);
            var readOptions = new EmailStoreItemReadOptions(
                EmailStoreItemReadParts.All, preferStreamingAttachmentContent: true);
            int sourceOrdinal = 0;
            foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
                cancellationToken.ThrowIfCancellationRequested();
                sourceOrdinal++;
                if (sourceOrdinal <= state.InspectedItems) continue;
                EmailStoreFolderInfo? sourceFolder = Folders.FirstOrDefault(item => item.Id == reference.FolderId);
                if (sourceFolder?.IsSearchFolder == true && !effective.IncludeSearchFolders) {
                    state.SkippedItems++;
                    state.InspectedItems++;
                    effective.Progress?.Report(new EmailStorePstMigrationProgress(
                        state.InspectedItems, state.ConvertedItems, state.SkippedItems,
                        effective.CheckpointPath));
                    CheckpointMigrationInterval(writer, mappings, state, diagnostics, effective);
                    continue;
                }
                if (!state.FolderMap.TryGetValue(reference.FolderId, out string? destinationFolder)) {
                    state.SkippedItems++;
                    state.InspectedItems++;
                    diagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_PST_CONVERT_FOLDER_UNMAPPED",
                        "An item was skipped because its source folder could not be mapped.",
                        EmailStoreDiagnosticSeverity.Error, reference.Id));
                    if (!effective.ContinueOnItemError) {
                        throw new InvalidDataException("A source item folder could not be mapped.");
                    }
                    effective.Progress?.Report(new EmailStorePstMigrationProgress(
                        state.InspectedItems, state.ConvertedItems, state.SkippedItems,
                        effective.CheckpointPath));
                    CheckpointMigrationInterval(writer, mappings, state, diagnostics, effective);
                    continue;
                }
                EmailStoreItem? item = null;
                try {
                    item = ReadItem(reference, readOptions, cancellationToken);
                } catch (Exception exception) when (effective.ContinueOnItemError &&
                    (exception is InvalidDataException || exception is NotSupportedException ||
                     exception is IOException || exception is EmailStoreLimitExceededException)) {
                    state.SkippedItems++;
                    diagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_PST_CONVERT_ITEM_SKIPPED",
                        string.Concat("A source item could not be read: ", exception.Message),
                        EmailStoreDiagnosticSeverity.Error, reference.Id));
                }
                if (item != null) {
                    destinationMutationUncertain = true;
                    string destinationItemId = writer.AddItem(destinationFolder, item.Document,
                        reference.IsAssociated, cancellationToken);
                    int conversionOrdinal = checked(state.ConvertedItems + 1);
                    if (effective.VerifyAfterWrite) {
                        mappings!.Add(conversionOrdinal, reference,
                            destinationFolder, destinationItemId);
                    }
                    state.ConvertedItems = conversionOrdinal;
                    if (reference.IsOrphaned) {
                        diagnostics.Add(new EmailStoreDiagnostic(
                            "EMAIL_STORE_PST_CONVERT_ORPHAN_RECOVERED",
                            "An item absent from its source contents table was recovered from the source index and copied.",
                            EmailStoreDiagnosticSeverity.Information, reference.Id));
                    }
                }
                state.InspectedItems++;
                destinationMutationUncertain = false;
                effective.Progress?.Report(new EmailStorePstMigrationProgress(
                    state.InspectedItems, state.ConvertedItems, state.SkippedItems,
                    effective.CheckpointPath));
                CheckpointMigrationInterval(writer, mappings, state, diagnostics, effective);
            }

            EnsureMigrationSourceUnchanged(sourceIdentity, cancellationToken);
            if (effective.FailOnDataLoss && diagnostics.Any(item =>
                item.Severity != EmailStoreDiagnosticSeverity.Information)) {
                throw new InvalidOperationException(
                    "Store conversion produced fidelity diagnostics and FailOnDataLoss is enabled.");
            }
            string completedWriterDestination = state.WriterDestination;
            completionStarted = true;
            EmailStorePstWriteReport writeReport = writer.Complete(cancellationToken);
            writerCompleted = true;
            diagnostics.AddRange(writeReport.Diagnostics);
            EmailStorePstVerificationReport? verification = effective.VerifyAfterWrite
                ? VerifyPstConversion(completedWriterDestination, mappings!, effective, diagnostics,
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
            mappings?.DeleteOnDispose();
            return new EmailStorePstConversionReport(Format, writeReport,
                Folders.Count, state.ConvertedItems, state.SkippedItems, verification,
                diagnostics.ToArray(), sourceIdentity, wasResumed);
        } catch {
            if (destinationMutationUncertain && !writerCompleted && writer != null) {
                writer.Abandon();
                mappings?.DeleteOnDispose();
            } else if (completionStarted && !writerCompleted && writer != null) {
                writer.Abandon();
                mappings?.DeleteOnDispose();
            } else if (!writerCompleted && writer != null && effective.CheckpointPath != null && stateReady) {
                if (effective.PartialResultPolicy == EmailStorePartialResultPolicy.RetainResumableState) {
                    CommitMigrationCheckpoint(writer, mappings, state!, diagnostics);
                } else {
                    writer.Abandon();
                    mappings?.DeleteOnDispose();
                }
            } else if (!writerCompleted && writer != null && !stateReady) {
                if (!resumeRequested) {
                    writer.Abandon();
                    mappings?.DeleteOnDispose();
                }
            } else if (writerCompleted) {
                mappings?.DeleteOnDispose();
            }
            throw;
        } finally {
            mappings?.Dispose();
            writer?.Dispose();
            OfficeFileCommit.DeleteIfExists(stagingPath);
            OfficeFileCommit.DeleteIfExists(manifestStagingPath);
        }
    }

    private static void CheckpointMigrationInterval(EmailStorePstWriter writer,
        PstConversionMappingJournal? mappings, PstConversionCheckpointState state,
        IList<EmailStoreDiagnostic> diagnostics, EmailStorePstConversionOptions options) {
        if (options.CheckpointPath == null ||
            state.InspectedItems % options.CheckpointIntervalItems != 0) return;
        CommitMigrationCheckpoint(writer, mappings, state, diagnostics);
    }

    private static void CommitMigrationCheckpoint(EmailStorePstWriter writer,
        PstConversionMappingJournal? mappings, PstConversionCheckpointState state,
        IList<EmailStoreDiagnostic> diagnostics) {
        if (mappings != null) {
            state.MappingLength = mappings.FlushDurable();
            state.MappingCount = mappings.Count;
        }
        state.Diagnostics.Clear();
        state.Diagnostics.AddRange(diagnostics);
        writer.Checkpoint(state.Serialize());
    }

    private void EnsureMigrationSourceUnchanged(EmailStoreSourceIdentity expected,
        CancellationToken cancellationToken) {
        if (!StringComparer.Ordinal.Equals(expected.CatalogFingerprint,
                GetCatalogFingerprint(cancellationToken)) ||
            !StringComparer.Ordinal.Equals(expected.DurableFingerprint,
                GetDurableSourceFingerprint(cancellationToken))) {
            throw new InvalidDataException(
                "The email-store source changed while the migration was running.");
        }
    }

    private static void ValidateResumedMigration(PstConversionCheckpointState state,
        EmailStoreSourceIdentity source, string optionsFingerprint, string destination,
        EmailStorePstConversionOptions options) {
        if (state.SourceFormat != source.Format || state.SourceLength != source.Length ||
            !StringComparer.Ordinal.Equals(state.CatalogFingerprint, source.CatalogFingerprint) ||
            !StringComparer.Ordinal.Equals(state.DurableFingerprint, source.DurableFingerprint)) {
            throw new InvalidDataException(
                "The migration checkpoint belongs to a changed or different email-store source.");
        }
        if (!StringComparer.Ordinal.Equals(state.OptionsFingerprint, optionsFingerprint)) {
            throw new InvalidDataException(
                "The migration checkpoint was created with different destination or conversion options.");
        }
        if (!options.VerifyAfterWrite &&
            !EmailStorePathIdentity.AreEquivalent(state.WriterDestination, destination)) {
            throw new InvalidDataException("The migration checkpoint destination is inconsistent.");
        }
        if (options.VerifyAfterWrite) {
            if (!IsOwnedMigrationStagingPath(state.WriterDestination, destination)) {
                throw new InvalidDataException(
                    "The migration checkpoint staging destination is inconsistent.");
            }
            string expectedMapping = string.Concat(options.CheckpointPath, ".verify-map");
            if (state.MappingPath == null ||
                !EmailStorePathIdentity.AreEquivalent(state.MappingPath, expectedMapping)) {
                throw new InvalidDataException("The migration verification journal path is inconsistent.");
            }
        } else if (state.MappingPath != null || state.MappingCount != 0 || state.MappingLength != 0) {
            throw new InvalidDataException("The migration checkpoint contains an unexpected verification journal.");
        }
    }

    private static bool IsOwnedMigrationStagingPath(string stagingPath, string destinationPath) {
        string staging = Path.GetFullPath(stagingPath);
        string destination = Path.GetFullPath(destinationPath);
        if (EmailStorePathIdentity.AreEquivalent(staging, destination) ||
            !EmailStorePathIdentity.AreEquivalent(
                Path.GetDirectoryName(staging) ?? Directory.GetCurrentDirectory(),
                Path.GetDirectoryName(destination) ?? Directory.GetCurrentDirectory()) ||
            !string.Equals(Path.GetExtension(staging), Path.GetExtension(destination),
                StringComparison.OrdinalIgnoreCase)) return false;
        string name = Path.GetFileNameWithoutExtension(staging);
        const string prefix = ".officeimo-";
        return name.StartsWith(prefix, StringComparison.Ordinal) &&
            Guid.TryParseExact(name.Substring(prefix.Length), "N", out _);
    }

    private void ValidateMigrationCheckpointPath(string destination,
        EmailStorePstConversionOptions options) {
        if (options.CheckpointPath == null) return;
        ThrowIfStoreSourceDestination(options.CheckpointPath, "PST migration checkpoint");
        if (EmailStorePathIdentity.AreEquivalent(options.CheckpointPath, destination) ||
            (options.VerificationManifestPath != null && EmailStorePathIdentity.AreEquivalent(
                options.CheckpointPath, options.VerificationManifestPath))) {
            throw new InvalidOperationException(
                "The migration checkpoint must use a path distinct from source, destination, and manifest.");
        }
    }

    private void ValidateVerificationManifestPath(string destination,
        EmailStorePstConversionOptions options) {
        if (options.VerificationManifestPath == null) return;
        string manifest = Path.GetFullPath(options.VerificationManifestPath);
        ThrowIfStoreSourceDestination(manifest, "PST verification manifest");
        if (EmailStorePathIdentity.AreEquivalent(manifest, destination)) {
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
