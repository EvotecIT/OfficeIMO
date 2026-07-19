namespace OfficeIMO.Email.Store;

/// <summary>Outcome of merging multiple stores into a new Unicode PST.</summary>
public sealed class EmailStorePstMergeReport {
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
}
