namespace OfficeIMO.Email.Store;

/// <summary>
/// Resumable item-enumeration position. Reuse it only with the same store snapshot, folder scope, and metadata filter.
/// </summary>
public sealed class EmailStoreContentSearchCheckpoint {
    /// <summary>Creates a checkpoint at a zero-based item-enumeration offset.</summary>
    public EmailStoreContentSearchCheckpoint(long itemOffset) {
        if (itemOffset < 0 || itemOffset > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(itemOffset));
        ItemOffset = itemOffset;
    }

    /// <summary>Number of item references already processed in the selected enumeration scope.</summary>
    public long ItemOffset { get; }
}
