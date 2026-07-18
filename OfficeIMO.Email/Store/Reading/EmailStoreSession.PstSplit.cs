using OfficeIMO.Drawing.Internal;
using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Builds a read-only query/size partition plan. No item payloads are read and no files are written.
    /// </summary>
    public EmailStorePstSplitPlan PlanPstSplit(string outputBasePath,
        EmailStorePstSplitOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(outputBasePath)) {
            throw new ArgumentException("An output base path is required.", nameof(outputBasePath));
        }
        ThrowIfDisposed();
        EmailStorePstSplitOptions effective = options ?? new EmailStorePstSplitOptions();
        string outputBase = Path.GetFullPath(outputBasePath);
        EmailStoreTableQuery sourceQuery = effective.Query ?? new EmailStoreTableQuery(
            includeAssociatedItems: true,
            includeOrphanedItems: true,
            sorts: new[] {
                EmailStoreFields.ReceivedAt.Ascending(),
                EmailStoreFields.SentAt.Ascending()
            },
            maxItemsScanned: 1_000_000,
            pageSize: 1000);
        var query = new EmailStoreTableQuery(
            sourceQuery.FolderId,
            sourceQuery.IncludeDescendants,
            sourceQuery.IncludeAssociatedItems,
            sourceQuery.IncludeOrphanedItems,
            sourceQuery.Filter,
            sourceQuery.Sorts,
            sourceQuery.Projection,
            continuationToken: null,
            sourceQuery.MaxItemsScanned,
            pageSize: int.MaxValue);
        EmailStoreTablePage page = SearchPage(query, cancellationToken);
        var diagnostics = new List<EmailStoreDiagnostic>();
        if (page.ScanLimitReached) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_SPLIT_QUERY_LIMIT",
                "The query scan bound was reached; execution is blocked because the selected set is incomplete.",
                EmailStoreDiagnosticSeverity.Error,
                outputBase));
        }

        var parts = new List<EmailStorePstSplitPlanPart>();
        var current = new List<EmailStoreItemReference>();
        long currentBytes = 0;
        bool currentOversized = false;
        int unknownSizeItems = 0;
        foreach (EmailStoreTableRow row in page.Rows) {
            cancellationToken.ThrowIfCancellationRequested();
            long itemBytes = EstimateSplitItemBytes(row.Summary, effective, out bool unknown);
            if (unknown) unknownSizeItems++;
            if (current.Count > 0 && currentBytes > effective.MaxEstimatedBytesPerPart -
                Math.Min(itemBytes, effective.MaxEstimatedBytesPerPart)) {
                if (!CompletePlannedPart(parts, current, currentBytes, currentOversized,
                    outputBase, effective, diagnostics)) break;
                current = new List<EmailStoreItemReference>();
                currentBytes = 0;
                currentOversized = false;
            }
            if (parts.Count >= effective.MaxParts) break;
            current.Add(row.Reference);
            currentBytes = checked(currentBytes + itemBytes);
            if (itemBytes > effective.MaxEstimatedBytesPerPart) currentOversized = true;
        }
        if (current.Count > 0 && parts.Count < effective.MaxParts) {
            CompletePlannedPart(parts, current, currentBytes, currentOversized,
                outputBase, effective, diagnostics);
        }
        int selected = parts.Sum(part => part.Items.Count);
        if (selected < page.Rows.Count) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_SPLIT_PART_LIMIT",
                string.Concat(page.Rows.Count - selected,
                    " selected item(s) were omitted after MaxParts=",
                    effective.MaxParts.ToString(CultureInfo.InvariantCulture), "."),
                EmailStoreDiagnosticSeverity.Error,
                outputBase));
        }
        foreach (EmailStorePstSplitPlanPart part in parts) {
            ThrowIfStoreSourceDestination(part.DestinationPath, "PST split");
            if (File.Exists(part.DestinationPath) && !effective.OverwriteExisting) {
                diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_PST_SPLIT_DESTINATION_EXISTS",
                    "A planned part already exists and overwriteExisting is false.",
                    EmailStoreDiagnosticSeverity.Error,
                    part.DestinationPath));
            }
        }
        return new EmailStorePstSplitPlan(this, outputBase, effective, query.Explain(),
            page.ItemsScanned, page.Rows.Count, page.ScanLimitReached, unknownSizeItems,
            parts.AsReadOnly(), diagnostics.AsReadOnly());
    }

    /// <summary>
    /// Executes a dry-run plan by staging every part, reopening and semantically verifying every written item,
    /// then committing the complete output set with rollback protection for replaced destinations.
    /// </summary>
    public EmailStorePstSplitReport ExecutePstSplit(EmailStorePstSplitPlan plan,
        CancellationToken cancellationToken = default) {
        if (plan == null) throw new ArgumentNullException(nameof(plan));
        ThrowIfDisposed();
        plan.ValidateOwner(this);
        if (!plan.IsExecutable) throw new InvalidOperationException(
            "The PST split plan is incomplete or has blocking diagnostics and cannot be executed.");
        foreach (EmailStorePstSplitPlanPart part in plan.Parts) {
            cancellationToken.ThrowIfCancellationRequested();
            ThrowIfStoreSourceDestination(part.DestinationPath, "PST split");
            OfficeFileCommit.EnsureTargetDirectory(part.DestinationPath);
            if (File.Exists(part.DestinationPath) && !plan.Options.OverwriteExisting) {
                throw new IOException(string.Concat("The planned split destination now exists: ",
                    part.DestinationPath));
            }
        }

        var staged = new List<StagedSplitPart>(plan.Parts.Count);
        try {
            foreach (EmailStorePstSplitPlanPart part in plan.Parts) {
                cancellationToken.ThrowIfCancellationRequested();
                staged.Add(WriteAndVerifySplitPart(part, plan.Options, cancellationToken));
            }
            CommitSplitSet(staged, plan.Options.OverwriteExisting);
            var reports = staged.Select(item => item.ToCommittedReport()).ToArray();
            return new EmailStorePstSplitReport(plan, reports,
                plan.Diagnostics.Concat(reports.SelectMany(report => report.Diagnostics)).ToArray());
        } finally {
            foreach (StagedSplitPart part in staged) {
                OfficeFileCommit.DeleteIfExists(part.StagingPath);
                OfficeFileCommit.DeleteIfExists(part.BackupPath);
            }
        }
    }

    /// <summary>Plans and executes a verified split in one call.</summary>
    public EmailStorePstSplitReport SplitToPst(string outputBasePath,
        EmailStorePstSplitOptions? options = null,
        CancellationToken cancellationToken = default) =>
        ExecutePstSplit(PlanPstSplit(outputBasePath, options, cancellationToken), cancellationToken);

    private StagedSplitPart WriteAndVerifySplitPart(EmailStorePstSplitPlanPart part,
        EmailStorePstSplitOptions options, CancellationToken cancellationToken) {
        string stagingPath = OfficeFileCommit.CreateStagingPath(part.DestinationPath);
        var diagnostics = new List<EmailStoreDiagnostic>();
        int skipped = 0;
        var conversionOptions = new EmailStorePstConversionOptions(
            overwriteExisting: false,
            failOnDataLoss: options.FailOnDataLoss,
            continueOnItemError: options.ContinueOnItemError,
            includeAssociatedItems: true,
            includeOrphanedItems: true,
            includeSearchFolders: options.IncludeSearchFolders,
            maxItems: Math.Max(1, part.Items.Count),
            maxNestedMessageDepth: options.MaxNestedMessageDepth,
            displayName: string.Concat(DisplayName ?? "OfficeIMO", " - Part ",
                part.Number.ToString(CultureInfo.InvariantCulture)),
            verifyAfterWrite: true,
            verificationOptions: options.VerificationOptions,
            maxVerificationIssues: options.MaxVerificationIssues);
        var writerOptions = new EmailStorePstWriterOptions(
            conversionOptions.DisplayName,
            overwriteExisting: false,
            failOnDataLoss: options.FailOnDataLoss,
            maxFolderCount: Math.Max(1, Folders.Count + 8),
            maxItemCount: Math.Max(1, part.Items.Count),
            maxNestedMessageDepth: options.MaxNestedMessageDepth);
        try {
            EmailStorePstWriteReport writeReport;
            EmailStorePstVerificationReport verification;
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(stagingPath, writerOptions))
            using (var mappings = new PstConversionMappingJournal(stagingPath)) {
                Dictionary<string, string> folderMap = CreatePstFolderMap(
                    writer, conversionOptions, diagnostics);
                var readOptions = new EmailStoreItemReadOptions(
                    EmailStoreItemReadParts.All,
                    preferStreamingAttachmentContent: true);
                int written = 0;
                foreach (EmailStoreItemReference reference in part.Items) {
                    cancellationToken.ThrowIfCancellationRequested();
                    EmailStoreFolderInfo? sourceFolder = Folders.FirstOrDefault(folder =>
                        string.Equals(folder.Id, reference.FolderId, StringComparison.Ordinal));
                    if (sourceFolder?.IsSearchFolder == true && !options.IncludeSearchFolders) {
                        skipped++;
                        diagnostics.Add(new EmailStoreDiagnostic(
                            "EMAIL_STORE_PST_SPLIT_SEARCH_FOLDER_SKIPPED",
                            "A selected search-folder item was skipped because IncludeSearchFolders is false.",
                            EmailStoreDiagnosticSeverity.Warning,
                            reference.Id));
                        continue;
                    }
                    if (!folderMap.TryGetValue(reference.FolderId, out string? destinationFolder)) {
                        skipped++;
                        diagnostics.Add(new EmailStoreDiagnostic(
                            "EMAIL_STORE_PST_SPLIT_FOLDER_UNMAPPED",
                            "A selected item's folder was not mapped into the output part.",
                            EmailStoreDiagnosticSeverity.Error,
                            reference.Id));
                        if (!options.ContinueOnItemError) throw new InvalidDataException(
                            "A selected item folder could not be mapped.");
                        continue;
                    }
                    try {
                        EmailStoreItem item = ReadItem(reference, readOptions, cancellationToken);
                        string destinationItemId = writer.AddItem(destinationFolder,
                            item.Document, reference.IsAssociated, cancellationToken);
                        mappings.Add(++written, reference, destinationFolder, destinationItemId);
                    } catch (Exception exception) when (options.ContinueOnItemError &&
                        (exception is InvalidDataException || exception is NotSupportedException ||
                         exception is IOException || exception is EmailStoreLimitExceededException)) {
                        skipped++;
                        diagnostics.Add(new EmailStoreDiagnostic(
                            "EMAIL_STORE_PST_SPLIT_ITEM_SKIPPED",
                            string.Concat("A selected item could not be copied: ", exception.Message),
                            EmailStoreDiagnosticSeverity.Error,
                            reference.Id));
                    }
                }
                if (written == 0) throw new InvalidDataException(
                    "A planned split part did not contain any writable selected items.");
                writeReport = writer.Complete(cancellationToken);
                diagnostics.AddRange(writeReport.Diagnostics);
                verification = VerifyPstConversion(stagingPath, mappings,
                    conversionOptions, diagnostics, manifestStagingPath: null, cancellationToken);
            }
            if (!verification.IsSuccessful) throw new InvalidDataException(
                "A staged PST split part failed semantic verification and was not committed.");
            if (options.FailOnDataLoss && diagnostics.Any(diagnostic =>
                diagnostic.Severity != EmailStoreDiagnosticSeverity.Information)) {
                throw new InvalidDataException(
                    "A staged PST split part emitted a fidelity diagnostic and was not committed.");
            }
            return new StagedSplitPart(part, stagingPath, writeReport, verification,
                skipped, diagnostics.AsReadOnly());
        } catch {
            OfficeFileCommit.DeleteIfExists(stagingPath);
            throw;
        }
    }

    private static void CommitSplitSet(IReadOnlyList<StagedSplitPart> staged,
        bool overwriteExisting) {
        foreach (StagedSplitPart part in staged) {
            part.DestinationExisted = File.Exists(part.Plan.DestinationPath);
            if (part.DestinationExisted) {
                if (!overwriteExisting) throw new IOException(string.Concat(
                    "A split destination was created before commit: ", part.Plan.DestinationPath));
                part.BackupPath = OfficeFileCommit.CreateStagingPath(part.Plan.DestinationPath);
                File.Copy(part.Plan.DestinationPath, part.BackupPath, overwrite: false);
            }
        }

        var committed = new List<StagedSplitPart>();
        try {
            foreach (StagedSplitPart part in staged) {
                OfficeFileCommit.CommitTemporaryFile(part.StagingPath, part.Plan.DestinationPath,
                    part.DestinationExisted
                        ? OfficeFileCommit.ConflictPolicy.Replace
                        : OfficeFileCommit.ConflictPolicy.FailIfExists);
                part.StagingPath = string.Empty;
                committed.Add(part);
            }
            foreach (StagedSplitPart part in staged) {
                OfficeFileCommit.DeleteIfExists(part.BackupPath);
                part.BackupPath = null;
            }
        } catch (Exception commitException) {
            var rollbackFailures = new List<Exception>();
            for (int index = committed.Count - 1; index >= 0; index--) {
                StagedSplitPart part = committed[index];
                try {
                    if (part.DestinationExisted && part.BackupPath != null) {
                        OfficeFileCommit.CommitTemporaryFile(part.BackupPath,
                            part.Plan.DestinationPath, OfficeFileCommit.ConflictPolicy.Replace);
                        part.BackupPath = null;
                    } else {
                        OfficeFileCommit.DeleteIfExists(part.Plan.DestinationPath);
                    }
                } catch (Exception rollbackException) when (
                    rollbackException is IOException || rollbackException is UnauthorizedAccessException) {
                    rollbackFailures.Add(rollbackException);
                }
            }
            if (rollbackFailures.Count > 0) {
                rollbackFailures.Insert(0, commitException);
                throw new AggregateException(
                    "The PST split commit failed and at least one destination rollback also failed.",
                    rollbackFailures);
            }
            throw;
        }
    }

    private static long EstimateSplitItemBytes(EmailStoreItemSummary summary,
        EmailStorePstSplitOptions options, out bool unknown) {
        unknown = !summary.DeclaredSize.HasValue || summary.DeclaredSize.Value <= 0;
        long payload = unknown ? options.UnknownItemEstimateBytes : summary.DeclaredSize!.Value;
        return checked(payload + options.PerItemOverheadBytes);
    }

    private static bool CompletePlannedPart(ICollection<EmailStorePstSplitPlanPart> parts,
        IReadOnlyList<EmailStoreItemReference> items, long estimatedBytes, bool oversized,
        string outputBase, EmailStorePstSplitOptions options,
        ICollection<EmailStoreDiagnostic> diagnostics) {
        if (parts.Count >= options.MaxParts) return false;
        int number = parts.Count + 1;
        string destination = BuildSplitPartPath(outputBase, number);
        parts.Add(new EmailStorePstSplitPlanPart(number, destination, items.ToArray(),
            estimatedBytes, options.MaxEstimatedBytesPerPart, oversized));
        if (oversized) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_SPLIT_OVERSIZED_ITEM",
                "One selected item alone exceeds the configured estimated part target; it remains isolated in its own part.",
                EmailStoreDiagnosticSeverity.Warning,
                destination));
        }
        return true;
    }

    private static string BuildSplitPartPath(string outputBase, int number) {
        string directory = Path.GetDirectoryName(outputBase) ?? Path.GetFullPath(".");
        string extension = Path.GetExtension(outputBase);
        string stem = string.Equals(extension, ".pst", StringComparison.OrdinalIgnoreCase)
            ? Path.GetFileNameWithoutExtension(outputBase)
            : Path.GetFileName(outputBase);
        return Path.Combine(directory, string.Concat(stem, ".part",
            number.ToString("D3", CultureInfo.InvariantCulture), ".pst"));
    }

    private sealed class StagedSplitPart {
        internal StagedSplitPart(EmailStorePstSplitPlanPart plan, string stagingPath,
            EmailStorePstWriteReport writeReport, EmailStorePstVerificationReport verification,
            int skippedItems, IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
            Plan = plan;
            StagingPath = stagingPath;
            WriteReport = writeReport;
            Verification = verification;
            SkippedItems = skippedItems;
            Diagnostics = diagnostics;
        }
        internal EmailStorePstSplitPlanPart Plan { get; }
        internal string StagingPath { get; set; }
        internal string? BackupPath { get; set; }
        internal bool DestinationExisted { get; set; }
        internal EmailStorePstWriteReport WriteReport { get; }
        internal EmailStorePstVerificationReport Verification { get; }
        internal int SkippedItems { get; }
        internal IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
        internal EmailStorePstSplitPartReport ToCommittedReport() {
            var report = new EmailStorePstWriteReport(
                Plan.DestinationPath,
                WriteReport.FolderCount,
                WriteReport.ItemCount,
                new FileInfo(Plan.DestinationPath).Length,
                WriteReport.Diagnostics,
                WriteReport.DiagnosticsTruncated);
            return new EmailStorePstSplitPartReport(Plan, report, Verification,
                SkippedItems, Diagnostics);
        }
    }
}
