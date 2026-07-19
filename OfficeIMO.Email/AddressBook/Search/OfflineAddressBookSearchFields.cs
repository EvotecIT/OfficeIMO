namespace OfficeIMO.Email.AddressBook;

/// <summary>Semantic fields considered by bounded Offline Address Book search.</summary>
[Flags]
public enum OfflineAddressBookSearchFields {
    /// <summary>No fields.</summary>
    None = 0,
    /// <summary>Display, given, surname, initials, and account names.</summary>
    Names = 1,
    /// <summary>SMTP, X500, target, and proxy addresses.</summary>
    Addresses = 2,
    /// <summary>Company, department, title, and office.</summary>
    Organization = 4,
    /// <summary>Telephone and fax values.</summary>
    Phones = 8,
    /// <summary>Street, locality, region, postal code, and country.</summary>
    PostalAddress = 16,
    /// <summary>Directory comment.</summary>
    Comment = 32,
    /// <summary>Distribution-list member and member-of distinguished names.</summary>
    Membership = 64,
    /// <summary>All supported semantic fields.</summary>
    All = Names | Addresses | Organization | Phones | PostalAddress | Comment | Membership
}
