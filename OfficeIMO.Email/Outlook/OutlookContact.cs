namespace OfficeIMO.Email;

/// <summary>A structured postal address on an Outlook contact.</summary>
public sealed class OutlookPostalAddress {
    /// <summary>Outlook-composed multiline address.</summary>
    public string? Formatted { get; set; }
    /// <summary>Street address.</summary>
    public string? Street { get; set; }
    /// <summary>City or locality.</summary>
    public string? City { get; set; }
    /// <summary>State or province.</summary>
    public string? StateOrProvince { get; set; }
    /// <summary>Postal code.</summary>
    public string? PostalCode { get; set; }
    /// <summary>Country or region.</summary>
    public string? Country { get; set; }
    /// <summary>Post-office box.</summary>
    public string? PostOfficeBox { get; set; }
    /// <summary>Country code when Outlook supplies one.</summary>
    public string? CountryCode { get; set; }
}

/// <summary>An electronic address slot on an Outlook contact.</summary>
public sealed class OutlookContactEmailAddress {
    /// <summary>Address value.</summary>
    public string? Address { get; set; }
    /// <summary>Display name.</summary>
    public string? DisplayName { get; set; }
    /// <summary>Original display name.</summary>
    public string? OriginalDisplayName { get; set; }
    /// <summary>Address type such as SMTP, EX, or FAX.</summary>
    public string? AddressType { get; set; }
    /// <summary>Original MAPI entry identifier.</summary>
    public byte[]? OriginalEntryId { get; set; }
}

/// <summary>Telephone and fax fields on an Outlook contact.</summary>
public sealed class OutlookContactPhones {
    /// <summary>Primary business telephone.</summary>
    public string? Business { get; set; }
    /// <summary>Secondary business telephone.</summary>
    public string? Business2 { get; set; }
    /// <summary>Primary home telephone.</summary>
    public string? Home { get; set; }
    /// <summary>Secondary home telephone.</summary>
    public string? Home2 { get; set; }
    /// <summary>Mobile telephone.</summary>
    public string? Mobile { get; set; }
    /// <summary>Other telephone.</summary>
    public string? Other { get; set; }
    /// <summary>Primary telephone.</summary>
    public string? Primary { get; set; }
    /// <summary>Business fax.</summary>
    public string? BusinessFax { get; set; }
    /// <summary>Home fax.</summary>
    public string? HomeFax { get; set; }
    /// <summary>Primary fax.</summary>
    public string? PrimaryFax { get; set; }
    /// <summary>Assistant telephone.</summary>
    public string? Assistant { get; set; }
    /// <summary>Company main telephone.</summary>
    public string? CompanyMain { get; set; }
    /// <summary>Car telephone.</summary>
    public string? Car { get; set; }
    /// <summary>Radio telephone.</summary>
    public string? Radio { get; set; }
    /// <summary>Pager or beeper.</summary>
    public string? Pager { get; set; }
    /// <summary>Callback telephone.</summary>
    public string? Callback { get; set; }
    /// <summary>Telex number.</summary>
    public string? Telex { get; set; }
    /// <summary>Text telephone or TTY/TDD number.</summary>
    public string? TextTelephone { get; set; }
    /// <summary>ISDN number.</summary>
    public string? Isdn { get; set; }
}

/// <summary>Typed Outlook contact fields.</summary>
public sealed class OutlookContact {
    private readonly List<string> _children = new List<string>();
    /// <summary>Contact display name.</summary>
    public string? DisplayName { get; set; }
    /// <summary>Name prefix or title.</summary>
    public string? Prefix { get; set; }
    /// <summary>Initials.</summary>
    public string? Initials { get; set; }
    /// <summary>Given name.</summary>
    public string? GivenName { get; set; }
    /// <summary>Middle name.</summary>
    public string? MiddleName { get; set; }
    /// <summary>Surname.</summary>
    public string? Surname { get; set; }
    /// <summary>Generational suffix.</summary>
    public string? Generation { get; set; }
    /// <summary>Company name.</summary>
    public string? CompanyName { get; set; }
    /// <summary>Job title or function.</summary>
    public string? JobTitle { get; set; }
    /// <summary>Department.</summary>
    public string? Department { get; set; }
    /// <summary>File-as display value.</summary>
    public string? FileAs { get; set; }
    /// <summary>Nickname.</summary>
    public string? NickName { get; set; }
    /// <summary>Manager name.</summary>
    public string? ManagerName { get; set; }
    /// <summary>Assistant name.</summary>
    public string? AssistantName { get; set; }
    /// <summary>Spouse name.</summary>
    public string? SpouseName { get; set; }
    /// <summary>Children names.</summary>
    public IList<string> Children => _children;
    /// <summary>Profession.</summary>
    public string? Profession { get; set; }
    /// <summary>Preferred language.</summary>
    public string? Language { get; set; }
    /// <summary>General contact location.</summary>
    public string? Location { get; set; }
    /// <summary>Office location.</summary>
    public string? OfficeLocation { get; set; }
    /// <summary>Birthday.</summary>
    public DateTimeOffset? Birthday { get; set; }
    /// <summary>Wedding anniversary.</summary>
    public DateTimeOffset? WeddingAnniversary { get; set; }
    /// <summary>Whether Outlook marks the contact private.</summary>
    public bool? IsPrivate { get; set; }
    /// <summary>Whether the contact has a picture attachment.</summary>
    public bool? HasPicture { get; set; }
    /// <summary>Business address.</summary>
    public OutlookPostalAddress BusinessAddress { get; } = new OutlookPostalAddress();
    /// <summary>Home address.</summary>
    public OutlookPostalAddress HomeAddress { get; } = new OutlookPostalAddress();
    /// <summary>Other address.</summary>
    public OutlookPostalAddress OtherAddress { get; } = new OutlookPostalAddress();
    /// <summary>Named work-address fields used by newer Outlook versions.</summary>
    public OutlookPostalAddress WorkAddress { get; } = new OutlookPostalAddress();
    /// <summary>Telephone and fax fields.</summary>
    public OutlookContactPhones Phones { get; } = new OutlookContactPhones();
    /// <summary>First electronic-mail address.</summary>
    public OutlookContactEmailAddress Email1 { get; } = new OutlookContactEmailAddress();
    /// <summary>Second electronic-mail address.</summary>
    public OutlookContactEmailAddress Email2 { get; } = new OutlookContactEmailAddress();
    /// <summary>Third electronic-mail address.</summary>
    public OutlookContactEmailAddress Email3 { get; } = new OutlookContactEmailAddress();
    /// <summary>Compatibility alias for the first electronic-mail address.</summary>
    public string? Email1Address { get => Email1.Address; set => Email1.Address = value; }
    /// <summary>Compatibility alias for the primary business telephone.</summary>
    public string? BusinessPhone { get => Phones.Business; set => Phones.Business = value; }
    /// <summary>Compatibility alias for the primary home telephone.</summary>
    public string? HomePhone { get => Phones.Home; set => Phones.Home = value; }
    /// <summary>Compatibility alias for the mobile telephone.</summary>
    public string? MobilePhone { get => Phones.Mobile; set => Phones.Mobile = value; }
    /// <summary>Instant-messaging address.</summary>
    public string? InstantMessagingAddress { get; set; }
    /// <summary>Business home page.</summary>
    public string? BusinessHomePage { get; set; }
    /// <summary>Personal home page.</summary>
    public string? PersonalHomePage { get; set; }
    /// <summary>Contact HTML or homepage content.</summary>
    public string? Html { get; set; }
}
