using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

/// <summary>Results and completion state for one bounded search batch.</summary>
public sealed class OfflineAddressBookSearchReport {
    internal OfflineAddressBookSearchReport(
        IReadOnlyList<OfflineAddressBookSearchResult> results,
        IReadOnlyList<EmailDiagnostic> diagnostics,
        int entriesScanned,
        int entriesSkipped,
        bool stoppedAtEntryLimit,
        bool stoppedAtResultLimit,
        OfflineAddressBookSearchCheckpoint? nextCheckpoint) {
        Results = results;
        Diagnostics = diagnostics;
        EntriesScanned = entriesScanned;
        EntriesSkipped = entriesSkipped;
        StoppedAtEntryLimit = stoppedAtEntryLimit;
        StoppedAtResultLimit = stoppedAtResultLimit;
        NextCheckpoint = nextCheckpoint;
    }

    /// <summary>Matches collected in this batch.</summary>
    public IReadOnlyList<OfflineAddressBookSearchResult> Results { get; }
    /// <summary>Search-specific diagnostics. Session diagnostics remain on the session.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>Records processed in this batch.</summary>
    public int EntriesScanned { get; }
    /// <summary>Corrupt or over-limit records skipped in this batch.</summary>
    public int EntriesSkipped { get; }
    /// <summary>Whether another record existed after the scan bound was reached.</summary>
    public bool StoppedAtEntryLimit { get; }
    /// <summary>Whether another record existed after the result bound was reached.</summary>
    public bool StoppedAtResultLimit { get; }
    /// <summary>Exact-position checkpoint for the next batch, or null when the selected scope was exhausted.</summary>
    public OfflineAddressBookSearchCheckpoint? NextCheckpoint { get; }
    /// <summary>Whether the selected address-list scope was exhausted.</summary>
    public bool IsComplete => NextCheckpoint == null;
}
