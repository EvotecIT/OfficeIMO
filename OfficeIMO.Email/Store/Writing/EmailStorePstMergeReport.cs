namespace OfficeIMO.Email.Store;

/// <summary>Outcome of merging multiple stores into a new Unicode PST.</summary>
public sealed class EmailStorePstMergeReport : IOfficeConversionReport {
    private readonly IReadOnlyList<OfficeConversionFidelityDiagnostic> _fidelityDiagnostics;

    internal EmailStorePstMergeReport(EmailStorePstWriteReport writeReport,
        IReadOnlyList<EmailStoreMergeSourceReport> sources, int inspectedItems,
        int writtenItems, int duplicateItems, int skippedItems, int retryCount,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics, bool diagnosticsTruncated) {
        WriteReport = writeReport;
        Sources = sources;
        InspectedItems = inspectedItems;
        WrittenItems = writtenItems;
        DuplicateItems = duplicateItems;
        SkippedItems = skippedItems;
        RetryCount = retryCount;
        Diagnostics = diagnostics;
        DiagnosticsTruncated = diagnosticsTruncated;
        var aggregate = new List<OfficeConversionFidelityDiagnostic>();
        if (duplicateItems > 0) {
            aggregate.Add(EmailStoreFidelityProjection.Create(
                "EMAIL_STORE_MERGE_DUPLICATES_OMITTED",
                duplicateItems + " semantically duplicate source item(s) were omitted by the selected merge policy.",
                OfficeConversionLossKind.Omission,
                writeReport.DestinationPath));
        }
        if (skippedItems > 0) {
            aggregate.Add(EmailStoreFidelityProjection.Create(
                "EMAIL_STORE_MERGE_ITEMS_OMITTED",
                skippedItems + " source item(s) were skipped during merge.",
                OfficeConversionLossKind.Omission,
                writeReport.DestinationPath));
        }
        if (sources.Any(static source => !source.Completed)) {
            aggregate.Add(EmailStoreFidelityProjection.Create(
                "EMAIL_STORE_MERGE_SOURCE_INCOMPLETE",
                "At least one merge source was not enumerated to completion.",
                OfficeConversionLossKind.Failure,
                writeReport.DestinationPath));
        }
        if (diagnosticsTruncated) {
            aggregate.Add(EmailStoreFidelityProjection.Create(
                "EMAIL_STORE_MERGE_DIAGNOSTICS_TRUNCATED",
                "Additional merge diagnostics exceeded the configured retention bound.",
                OfficeConversionLossKind.Failure,
                writeReport.DestinationPath));
        }
        _fidelityDiagnostics = EmailStoreFidelityProjection.Append(
            EmailStoreFidelityProjection.Project(diagnostics), aggregate.ToArray());
    }

    /// <summary>Committed destination PST report.</summary>
    public EmailStorePstWriteReport WriteReport { get; }
    /// <summary>Per-source aggregate reports.</summary>
    public IReadOnlyList<EmailStoreMergeSourceReport> Sources { get; }
    /// <summary>Total source items inspected.</summary>
    public int InspectedItems { get; }
    /// <summary>Total items written.</summary>
    public int WrittenItems { get; }
    /// <summary>Total semantic duplicates omitted.</summary>
    public int DuplicateItems { get; }
    /// <summary>Total items skipped after reported failures or mapping decisions.</summary>
    public int SkippedItems { get; }
    /// <summary>Total transient source I/O retries consumed.</summary>
    public int RetryCount { get; }
    /// <summary>Bounded merge and writer diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>Whether additional detailed diagnostics were omitted.</summary>
    public bool DiagnosticsTruncated { get; }
    /// <summary>Whether any warning or error was reported.</summary>
    public bool HasIssues => Diagnostics.Any(diagnostic =>
        diagnostic.Severity != EmailStoreDiagnosticSeverity.Information);

    /// <summary>Category-preserving merge, omission, completion, and writer diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;

    /// <summary>Whether merge approximated, omitted, or failed to preserve selected source content.</summary>
    public bool HasLoss => _fidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>Throws when merge reported possible content loss.</summary>
    public void RequireNoLoss() => EmailStoreFidelityProjection.RequireNoLoss(_fidelityDiagnostics);
}
