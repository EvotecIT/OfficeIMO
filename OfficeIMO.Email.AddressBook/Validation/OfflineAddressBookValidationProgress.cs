namespace OfficeIMO.Email.AddressBook;

/// <summary>Aggregate-only progress for an explicit validation pass.</summary>
public sealed class OfflineAddressBookValidationProgress {
    internal OfflineAddressBookValidationProgress(int addressListsCompleted,
        long bytesHashed, long entriesScanned, long entriesSkipped) {
        AddressListsCompleted = addressListsCompleted;
        BytesHashed = bytesHashed;
        EntriesScanned = entriesScanned;
        EntriesSkipped = entriesSkipped;
    }

    /// <summary>Selected address lists completed.</summary>
    public int AddressListsCompleted { get; }
    /// <summary>Payload bytes processed by checksum validation.</summary>
    public long BytesHashed { get; }
    /// <summary>Entry records walked.</summary>
    public long EntriesScanned { get; }
    /// <summary>Value-level records skipped after recoverable decode failures.</summary>
    public long EntriesSkipped { get; }
}
