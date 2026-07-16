namespace OfficeIMO.Email.AddressBook;

/// <summary>Stable record reference within one open OAB snapshot.</summary>
public sealed class OfflineAddressBookEntryReference {
    internal OfflineAddressBookEntryReference(string addressListId, int addressListIndex,
        long entryIndex, long recordOffset, int recordLength, Guid snapshotId) {
        AddressListId = addressListId;
        AddressListIndex = addressListIndex;
        EntryIndex = entryIndex;
        RecordOffset = recordOffset;
        RecordLength = recordLength;
        SnapshotId = snapshotId;
        Id = string.Concat(addressListId, ":", entryIndex.ToString("D10", CultureInfo.InvariantCulture));
    }

    /// <summary>Stable reference text for the current session snapshot.</summary>
    public string Id { get; }
    /// <summary>Owning address-list identifier.</summary>
    public string AddressListId { get; }
    /// <summary>Zero-based address-list index.</summary>
    public int AddressListIndex { get; }
    /// <summary>Zero-based record index within the address list.</summary>
    public long EntryIndex { get; }
    /// <summary>Record offset relative to the beginning of the Full Details component.</summary>
    public long RecordOffset { get; }
    /// <summary>Encoded record length.</summary>
    public int RecordLength { get; }
    internal Guid SnapshotId { get; }
}
