namespace OfficeIMO.Email.AddressBook;

/// <summary>Bounds and selection for an explicit integrity pass.</summary>
public sealed class OfflineAddressBookValidationOptions {
    /// <summary>Creates validation options.</summary>
    public OfflineAddressBookValidationOptions(
        OfflineAddressBookValidationMode mode = OfflineAddressBookValidationMode.FullDecode,
        bool validateChecksum = true,
        string? addressListId = null,
        int maxEntriesPerAddressList = int.MaxValue,
        long maxChecksumBytesPerFile = 64L * 1024 * 1024 * 1024,
        int progressEntryInterval = 10_000,
        long progressByteInterval = 64L * 1024 * 1024,
        bool continueOnEntryError = true) {
        if (!Enum.IsDefined(typeof(OfflineAddressBookValidationMode), mode)) {
            throw new ArgumentOutOfRangeException(nameof(mode));
        }
        if (maxEntriesPerAddressList <= 0) throw new ArgumentOutOfRangeException(nameof(maxEntriesPerAddressList));
        if (maxChecksumBytesPerFile <= 0) throw new ArgumentOutOfRangeException(nameof(maxChecksumBytesPerFile));
        if (progressEntryInterval <= 0) throw new ArgumentOutOfRangeException(nameof(progressEntryInterval));
        if (progressByteInterval <= 0) throw new ArgumentOutOfRangeException(nameof(progressByteInterval));
        Mode = mode;
        ValidateChecksum = validateChecksum;
        AddressListId = string.IsNullOrWhiteSpace(addressListId) ? null : addressListId;
        MaxEntriesPerAddressList = maxEntriesPerAddressList;
        MaxChecksumBytesPerFile = maxChecksumBytesPerFile;
        ProgressEntryInterval = progressEntryInterval;
        ProgressByteInterval = progressByteInterval;
        ContinueOnEntryError = continueOnEntryError;
    }

    /// <summary>Validation depth.</summary>
    public OfflineAddressBookValidationMode Mode { get; }
    /// <summary>Whether to recalculate and compare the OAB header CRC.</summary>
    public bool ValidateChecksum { get; }
    /// <summary>Optional address-list identifier.</summary>
    public string? AddressListId { get; }
    /// <summary>Maximum records walked in each selected address list.</summary>
    public int MaxEntriesPerAddressList { get; }
    /// <summary>Maximum bytes hashed in one Full Details component.</summary>
    public long MaxChecksumBytesPerFile { get; }
    /// <summary>Record interval between progress notifications.</summary>
    public int ProgressEntryInterval { get; }
    /// <summary>Byte interval between checksum progress notifications.</summary>
    public long ProgressByteInterval { get; }
    /// <summary>Whether a value-level decode error is diagnosed and skipped when framing remains trustworthy.</summary>
    public bool ContinueOnEntryError { get; }
}
