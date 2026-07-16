namespace OfficeIMO.Email.Store;

/// <summary>Results and completion state for one bounded content-search batch.</summary>
public sealed class EmailStoreContentSearchReport {
    internal EmailStoreContentSearchReport(
        IReadOnlyList<EmailStoreContentSearchResult> results,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics,
        int itemsScanned,
        int itemsSkipped,
        bool stoppedAtItemLimit,
        bool stoppedAtResultLimit,
        EmailStoreContentSearchCheckpoint? nextCheckpoint) {
        Results = results;
        Diagnostics = diagnostics;
        ItemsScanned = itemsScanned;
        ItemsSkipped = itemsSkipped;
        StoppedAtItemLimit = stoppedAtItemLimit;
        StoppedAtResultLimit = stoppedAtResultLimit;
        NextCheckpoint = nextCheckpoint;
    }

    /// <summary>Matches collected in this batch.</summary>
    public IReadOnlyList<EmailStoreContentSearchResult> Results { get; }
    /// <summary>Search-specific diagnostics. Store-open diagnostics remain on the session.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>Item references processed in this batch.</summary>
    public int ItemsScanned { get; }
    /// <summary>Corrupt or over-limit items skipped in this batch.</summary>
    public int ItemsSkipped { get; }
    /// <summary>Whether another item existed after the scan bound was reached.</summary>
    public bool StoppedAtItemLimit { get; }
    /// <summary>Whether another item existed after the result bound was reached.</summary>
    public bool StoppedAtResultLimit { get; }
    /// <summary>Checkpoint for the next batch, or null when the selected enumeration was exhausted.</summary>
    public EmailStoreContentSearchCheckpoint? NextCheckpoint { get; }
    /// <summary>Whether the selected enumeration was exhausted.</summary>
    public bool IsComplete => NextCheckpoint == null;
}
