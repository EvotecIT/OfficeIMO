namespace OfficeIMO.Email.AddressBook;

/// <summary>Aggregate-only progress for a bounded address-book search batch.</summary>
public sealed class OfflineAddressBookSearchProgress {
    internal OfflineAddressBookSearchProgress(int entriesScanned, int matches, int entriesSkipped) {
        EntriesScanned = entriesScanned;
        Matches = matches;
        EntriesSkipped = entriesSkipped;
    }

    /// <summary>Records processed in this batch.</summary>
    public int EntriesScanned { get; }
    /// <summary>Matches collected in this batch.</summary>
    public int Matches { get; }
    /// <summary>Corrupt or over-limit records skipped in this batch.</summary>
    public int EntriesSkipped { get; }
}
