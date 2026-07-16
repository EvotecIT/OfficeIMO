namespace OfficeIMO.Email.AddressBook;

/// <summary>Selection and error policy for lazy OAB enumeration.</summary>
public sealed class OfflineAddressBookEnumerationOptions {
    /// <summary>Creates enumeration options.</summary>
    public OfflineAddressBookEnumerationOptions(
        string? addressListId = null,
        int maxEntries = int.MaxValue,
        bool continueOnEntryError = true) {
        if (maxEntries <= 0) throw new ArgumentOutOfRangeException(nameof(maxEntries));
        AddressListId = string.IsNullOrWhiteSpace(addressListId) ? null : addressListId;
        MaxEntries = maxEntries;
        ContinueOnEntryError = continueOnEntryError;
    }

    /// <summary>Optional list identifier. Null enumerates every discovered Full Details file.</summary>
    public string? AddressListId { get; }
    /// <summary>Maximum successfully decoded entries returned.</summary>
    public int MaxEntries { get; }
    /// <summary>Whether a corrupt or over-limit record is diagnosed and skipped when its boundary is known.</summary>
    public bool ContinueOnEntryError { get; }
}
