namespace OfficeIMO.Email.AddressBook;

/// <summary>
/// Resumable position within one session snapshot. A checkpoint includes the exact next record offset, so resuming
/// a huge address list does not require rescanning earlier records.
/// </summary>
public sealed class OfflineAddressBookSearchCheckpoint {
    internal OfflineAddressBookSearchCheckpoint(string addressListId, int addressListIndex,
        long entryIndex, long recordOffset, Guid snapshotId) {
        AddressListId = addressListId;
        AddressListIndex = addressListIndex;
        EntryIndex = entryIndex;
        RecordOffset = recordOffset;
        SnapshotId = snapshotId;
    }

    /// <summary>Address-list identifier in the session snapshot.</summary>
    public string AddressListId { get; }
    /// <summary>Zero-based address-list index.</summary>
    public int AddressListIndex { get; }
    /// <summary>Zero-based index of the next record to scan.</summary>
    public long EntryIndex { get; }
    /// <summary>Exact next-record offset relative to its Full Details component.</summary>
    public long RecordOffset { get; }
    internal Guid SnapshotId { get; }
}
