namespace OfficeIMO.Email;

/// <summary>Represents an email or Outlook address.</summary>
public sealed class EmailAddress {
    /// <summary>Creates an address.</summary>
    public EmailAddress(string? address, string? displayName = null, string? rawValue = null) {
        Address = address;
        DisplayName = displayName;
        RawValue = rawValue;
    }

    /// <summary>SMTP or source-specific address value.</summary>
    public string? Address { get; set; }

    /// <summary>Human-readable display name.</summary>
    public string? DisplayName { get; set; }

    /// <summary>Original source spelling when available.</summary>
    public string? RawValue { get; set; }

    /// <summary>Source address type such as SMTP, EX, X400, or FAX.</summary>
    public string? AddressType { get; set; }

    /// <inheritdoc />
    public override string ToString() {
        if (!string.IsNullOrWhiteSpace(DisplayName) && !string.IsNullOrWhiteSpace(Address)) {
            return string.Concat(DisplayName, " <", Address, ">");
        }
        return DisplayName ?? Address ?? RawValue ?? string.Empty;
    }
}

/// <summary>Associates an address with its recipient role.</summary>
public sealed class EmailRecipient {
    private readonly List<MapiProperty> _mapiProperties = new List<MapiProperty>();
    private MapiPropertyBag? _mapi;
    /// <summary>Creates a recipient.</summary>
    public EmailRecipient(EmailRecipientKind kind, EmailAddress address) {
        Kind = kind;
        Address = address ?? throw new ArgumentNullException(nameof(address));
    }

    /// <summary>Recipient role.</summary>
    public EmailRecipientKind Kind { get; set; }

    /// <summary>Recipient address.</summary>
    public EmailAddress Address { get; set; }

    /// <summary>MAPI row identifier when the source supplies one.</summary>
    public int? MapiRowId { get; set; }

    /// <summary>MAPI object type, normally 6 for a messaging user.</summary>
    public int? MapiObjectType { get; set; }

    /// <summary>MAPI display type.</summary>
    public int? MapiDisplayType { get; set; }

    /// <summary>Extended MAPI display type used to distinguish rooms.</summary>
    public int? MapiDisplayTypeEx { get; set; }

    /// <summary>Recipient-level MAPI properties.</summary>
    public IList<MapiProperty> MapiProperties => _mapiProperties;

    /// <summary>Typed MAPI access backed by the exact <see cref="MapiProperties"/> collection.</summary>
    public MapiPropertyBag Mapi => _mapi ?? (_mapi = new MapiPropertyBag(_mapiProperties));
}
