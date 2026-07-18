namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Scans lightweight summaries and produces a no-write capacity/selection plan for a distinct compacted PST.
    /// </summary>
    public EmailStorePstCompactionPlan PlanPstCompaction(string destinationPath,
        EmailStorePstCompactionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(destinationPath)) {
            throw new ArgumentException("A destination path is required.", nameof(destinationPath));
        }
        ThrowIfDisposed();
        EmailStorePstCompactionOptions effective = options ??
            new EmailStorePstCompactionOptions();
        string destination = Path.GetFullPath(destinationPath);
        var diagnostics = new List<EmailStoreDiagnostic>();
        if (Format != EmailStoreFormat.Pst) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_COMPACT_SOURCE_FORMAT",
                "Verified PST compaction requires a PST source. Use ExportToPst for other store formats.",
                EmailStoreDiagnosticSeverity.Error,
                destination));
        }
        if (IsPstPasswordProtected) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_COMPACT_PASSWORD_UNSUPPORTED",
                "The managed writer cannot preserve legacy PST password protection, so compaction is blocked.",
                EmailStoreDiagnosticSeverity.Error,
                destination));
        }
        try {
            ThrowIfStoreSourceDestination(destination, "PST compaction");
        } catch (InvalidOperationException exception) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_COMPACT_DESTINATION_SOURCE",
                exception.Message,
                EmailStoreDiagnosticSeverity.Error,
                destination));
        }
        if (File.Exists(destination) && !effective.OverwriteExisting) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_COMPACT_DESTINATION_EXISTS",
                "The compacted destination already exists and overwriteExisting is false.",
                EmailStoreDiagnosticSeverity.Error,
                destination));
        }

        int probeLimit = effective.MaxItems == int.MaxValue ? int.MaxValue : effective.MaxItems + 1;
        var enumeration = new EmailStoreEnumerationOptions(
            includeAssociatedItems: effective.IncludeAssociatedItems,
            includeOrphanedItems: effective.IncludeOrphanedItems,
            maxItems: probeLimit);
        int scanned = 0;
        int selected = 0;
        int associated = 0;
        int orphaned = 0;
        int excludedSearch = 0;
        int unknown = 0;
        long estimate = effective.FixedPstOverheadBytes;
        bool limit = false;
        foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (scanned >= effective.MaxItems) { limit = true; break; }
            scanned++;
            EmailStoreFolderInfo? folder = Folders.FirstOrDefault(item =>
                string.Equals(item.Id, reference.FolderId, StringComparison.Ordinal));
            if (folder?.IsSearchFolder == true && !effective.IncludeSearchFolders) {
                excludedSearch++;
                continue;
            }
            EmailStoreItemSummary summary = ReadSummary(reference, cancellationToken);
            bool unknownSize = !summary.DeclaredSize.HasValue || summary.DeclaredSize.Value <= 0;
            long payload = unknownSize
                ? effective.UnknownItemEstimateBytes
                : summary.DeclaredSize!.Value;
            estimate = checked(estimate + payload + effective.PerItemOverheadBytes);
            if (unknownSize) unknown++;
            if (reference.IsAssociated) associated++;
            if (reference.IsOrphaned) orphaned++;
            selected++;
        }
        if (limit) diagnostics.Add(new EmailStoreDiagnostic(
            "EMAIL_STORE_PST_COMPACT_ITEM_LIMIT",
            string.Concat("At least one item was omitted after MaxItems=",
                effective.MaxItems.ToString(CultureInfo.InvariantCulture), "."),
            EmailStoreDiagnosticSeverity.Error,
            destination));
        if (excludedSearch > 0) diagnostics.Add(new EmailStoreDiagnostic(
            "EMAIL_STORE_PST_COMPACT_SEARCH_ITEMS_EXCLUDED",
            string.Concat(excludedSearch.ToString(CultureInfo.InvariantCulture),
                " search-folder item(s) are intentionally excluded by the compaction policy."),
            effective.FailOnDataLoss
                ? EmailStoreDiagnosticSeverity.Error
                : EmailStoreDiagnosticSeverity.Warning,
            destination));
        diagnostics.AddRange(Diagnostics);
        return new EmailStorePstCompactionPlan(destination, effective, SourceLength,
            scanned, selected, associated, orphaned, excludedSearch, unknown,
            estimate, limit, diagnostics.AsReadOnly());
    }

    /// <summary>
    /// Rewrites the selected live PST data to a distinct destination, reopens it, compares every written item,
    /// and commits only after verification. This method never replaces the open source PST.
    /// </summary>
    public EmailStorePstCompactionReport CompactToPst(string destinationPath,
        EmailStorePstCompactionOptions? options = null,
        CancellationToken cancellationToken = default) {
        EmailStorePstCompactionPlan plan = PlanPstCompaction(
            destinationPath, options, cancellationToken);
        if (!plan.IsExecutable) throw new InvalidOperationException(
            "The PST compaction plan has blocking diagnostics and cannot be executed.");
        EmailStorePstCompactionOptions effective = plan.Options;
        var conversionOptions = new EmailStorePstConversionOptions(
            overwriteExisting: effective.OverwriteExisting,
            failOnDataLoss: effective.FailOnDataLoss,
            continueOnItemError: effective.ContinueOnItemError,
            includeAssociatedItems: effective.IncludeAssociatedItems,
            includeOrphanedItems: effective.IncludeOrphanedItems,
            includeSearchFolders: effective.IncludeSearchFolders,
            maxItems: effective.MaxItems,
            maxNestedMessageDepth: effective.MaxNestedMessageDepth,
            displayName: effective.DisplayName,
            verifyAfterWrite: true,
            verificationOptions: effective.VerificationOptions,
            maxVerificationIssues: effective.MaxVerificationIssues);
        EmailStorePstConversionReport conversion = ExportToPst(
            plan.DestinationPath, conversionOptions, cancellationToken);
        return new EmailStorePstCompactionReport(plan, conversion);
    }
}
