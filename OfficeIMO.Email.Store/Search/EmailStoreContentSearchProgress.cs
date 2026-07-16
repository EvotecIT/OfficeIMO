namespace OfficeIMO.Email.Store;

/// <summary>Aggregate-only progress for a bounded content-search batch.</summary>
public sealed class EmailStoreContentSearchProgress {
    internal EmailStoreContentSearchProgress(int itemsScanned, int matches, int itemsSkipped) {
        ItemsScanned = itemsScanned;
        Matches = matches;
        ItemsSkipped = itemsSkipped;
    }

    /// <summary>Item references processed in this batch.</summary>
    public int ItemsScanned { get; }
    /// <summary>Matches collected in this batch.</summary>
    public int Matches { get; }
    /// <summary>Corrupt or over-limit items skipped in this batch.</summary>
    public int ItemsSkipped { get; }
}
