using OfficeIMO.Email.AddressBook;

namespace OfficeIMO.Reader.Email;

/// <summary>Options for bounded OAB entry projection through OfficeIMO.Reader.</summary>
public sealed class ReaderEmailAddressBookOptions {
    private int _maxEntries = 10_000;
    private int _maxMultiValueItems = 100;

    /// <summary>Gets or sets core OAB reader limits. Registrations capture a defensive copy.</summary>
    public OfflineAddressBookReaderOptions? AddressBookOptions { get; set; }

    /// <summary>Optional address-list identifier.</summary>
    public string? AddressListId { get; set; }

    /// <summary>
    /// Optional bounded core query. Reader projects matching references only; use this to avoid ingesting an entire
    /// large address book.
    /// </summary>
    public OfflineAddressBookSearchQuery? Query { get; set; }

    /// <summary>Maximum entries projected in one call. Default: 10,000.</summary>
    public int MaxEntries {
        get => _maxEntries;
        set {
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value));
            _maxEntries = value;
        }
    }

    /// <summary>Whether recoverable record failures produce diagnostic-only results and reading continues.</summary>
    public bool ContinueOnEntryError { get; set; } = true;

    /// <summary>
    /// Whether distribution-list member and member-of distinguished names are included in chunk text. Disabled by
    /// default because these lists can be large and are often unnecessary for general indexing.
    /// </summary>
    public bool IncludeMembershipValues { get; set; }

    /// <summary>Maximum emitted values from each proxy, phone, or membership collection. Default: 100.</summary>
    public int MaxMultiValueItems {
        get => _maxMultiValueItems;
        set {
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value));
            _maxMultiValueItems = value;
        }
    }

    /// <summary>
    /// Whether a complete source-file hash is calculated for native document results. Disabled by default so a
    /// selective read does not force an additional full pass over a huge OAB.
    /// </summary>
    public bool ComputeSourceHash { get; set; }
}
