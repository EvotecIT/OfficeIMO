using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

/// <summary>One schema-driven Offline Address Book object.</summary>
public sealed class OfflineAddressBookEntry {
    internal OfflineAddressBookEntry(
        OfflineAddressBookEntryReference reference,
        OfflineAddressBookListInfo addressList,
        IReadOnlyList<MapiProperty> properties,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        Reference = reference;
        AddressList = addressList;
        Properties = properties;
        Diagnostics = diagnostics;

        X500Address = OabPropertyValues.String(properties, OabPropertyTags.EmailAddress);
        SmtpAddress = OabPropertyValues.String(properties, OabPropertyTags.SmtpAddress);
        DisplayName = OabPropertyValues.String(properties, OabPropertyTags.DisplayName) ??
            OabPropertyValues.String(properties, OabPropertyTags.DisplayNamePrintable);
        Account = OabPropertyValues.String(properties, OabPropertyTags.Account);
        GivenName = OabPropertyValues.String(properties, OabPropertyTags.GivenName);
        Surname = OabPropertyValues.String(properties, OabPropertyTags.Surname);
        Initials = OabPropertyValues.String(properties, OabPropertyTags.Initials);
        CompanyName = OabPropertyValues.String(properties, OabPropertyTags.CompanyName);
        JobTitle = OabPropertyValues.String(properties, OabPropertyTags.Title);
        Department = OabPropertyValues.String(properties, OabPropertyTags.Department);
        OfficeLocation = OabPropertyValues.String(properties, OabPropertyTags.OfficeLocation);
        AssistantName = OabPropertyValues.String(properties, OabPropertyTags.Assistant);
        Comment = OabPropertyValues.String(properties, OabPropertyTags.Comment);
        TargetAddress = OabPropertyValues.String(properties, OabPropertyTags.TargetAddress);
        StreetAddress = OabPropertyValues.String(properties, OabPropertyTags.StreetAddress);
        Locality = OabPropertyValues.String(properties, OabPropertyTags.Locality);
        StateOrProvince = OabPropertyValues.String(properties, OabPropertyTags.StateOrProvince);
        PostalCode = OabPropertyValues.String(properties, OabPropertyTags.PostalCode);
        Country = OabPropertyValues.String(properties, OabPropertyTags.Country);
        BusinessTelephone = OabPropertyValues.String(properties, OabPropertyTags.BusinessTelephone);
        HomeTelephone = OabPropertyValues.String(properties, OabPropertyTags.HomeTelephone);
        MobileTelephone = OabPropertyValues.String(properties, OabPropertyTags.MobileTelephone);
        PrimaryFax = OabPropertyValues.String(properties, OabPropertyTags.PrimaryFax);
        AssistantTelephone = OabPropertyValues.String(properties, OabPropertyTags.AssistantTelephone);
        PagerTelephone = OabPropertyValues.String(properties, OabPropertyTags.PagerTelephone);
        BusinessTelephone2 = OabPropertyValues.Strings(properties, OabPropertyTags.BusinessTelephone2);
        HomeTelephone2 = OabPropertyValues.Strings(properties, OabPropertyTags.HomeTelephone2);
        ProxyAddresses = OabPropertyValues.Strings(properties, OabPropertyTags.ProxyAddresses);
        MemberDistinguishedNames = OabPropertyValues.Strings(properties, OabPropertyTags.Members);
        MemberOfDistinguishedNames = OabPropertyValues.Strings(properties, OabPropertyTags.MemberOf);
        TruncatedPropertyTags = OabPropertyValues.UInt32s(properties, OabPropertyTags.TruncatedProperties);
        RawObjectType = OabPropertyValues.UInt32(properties, OabPropertyTags.ObjectType);
        DisplayType = OabPropertyValues.UInt32(properties, OabPropertyTags.DisplayType);
        DisplayTypeEx = OabPropertyValues.UInt32(properties, OabPropertyTags.DisplayTypeEx);
        DistributionListMemberCount = OabPropertyValues.UInt32(properties, OabPropertyTags.DistributionListMemberCount);
        DistributionListExternalMemberCount = OabPropertyValues.UInt32(properties, OabPropertyTags.DistributionListExternalMemberCount);
        IsHierarchicalGroup = OabPropertyValues.Boolean(properties, OabPropertyTags.HierarchicalGroup);
        CanReceiveRichContent = OabPropertyValues.Boolean(properties, OabPropertyTags.SendRichInfo);
        ObjectType = RawObjectType.HasValue && Enum.IsDefined(typeof(OfflineAddressBookObjectType), (int)RawObjectType.Value)
            ? (OfflineAddressBookObjectType)RawObjectType.Value
            : OfflineAddressBookObjectType.Unknown;
    }

    /// <summary>Stable reference for this session snapshot.</summary>
    public OfflineAddressBookEntryReference Reference { get; }
    /// <summary>Owning address list.</summary>
    public OfflineAddressBookListInfo AddressList { get; }
    /// <summary>All decoded file-defined properties.</summary>
    public IReadOnlyList<MapiProperty> Properties { get; }
    /// <summary>Entry-scoped compatibility and fidelity diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>X500 distinguished name.</summary>
    public string? X500Address { get; }
    /// <summary>Primary SMTP address.</summary>
    public string? SmtpAddress { get; }
    /// <summary>Display name.</summary>
    public string? DisplayName { get; }
    /// <summary>Directory account name.</summary>
    public string? Account { get; }
    /// <summary>Given name.</summary>
    public string? GivenName { get; }
    /// <summary>Surname.</summary>
    public string? Surname { get; }
    /// <summary>Initials.</summary>
    public string? Initials { get; }
    /// <summary>Company name.</summary>
    public string? CompanyName { get; }
    /// <summary>Job title.</summary>
    public string? JobTitle { get; }
    /// <summary>Department.</summary>
    public string? Department { get; }
    /// <summary>Office location.</summary>
    public string? OfficeLocation { get; }
    /// <summary>Assistant name.</summary>
    public string? AssistantName { get; }
    /// <summary>Directory comment.</summary>
    public string? Comment { get; }
    /// <summary>Routing target address.</summary>
    public string? TargetAddress { get; }
    /// <summary>Street address.</summary>
    public string? StreetAddress { get; }
    /// <summary>City or locality.</summary>
    public string? Locality { get; }
    /// <summary>State or province.</summary>
    public string? StateOrProvince { get; }
    /// <summary>Postal code.</summary>
    public string? PostalCode { get; }
    /// <summary>Country or region.</summary>
    public string? Country { get; }
    /// <summary>Primary business telephone.</summary>
    public string? BusinessTelephone { get; }
    /// <summary>Primary home telephone.</summary>
    public string? HomeTelephone { get; }
    /// <summary>Mobile telephone.</summary>
    public string? MobileTelephone { get; }
    /// <summary>Primary fax number.</summary>
    public string? PrimaryFax { get; }
    /// <summary>Assistant telephone.</summary>
    public string? AssistantTelephone { get; }
    /// <summary>Pager telephone.</summary>
    public string? PagerTelephone { get; }
    /// <summary>Secondary business telephone numbers.</summary>
    public IReadOnlyList<string> BusinessTelephone2 { get; }
    /// <summary>Secondary home telephone numbers.</summary>
    public IReadOnlyList<string> HomeTelephone2 { get; }
    /// <summary>SMTP, X400, and other proxy addresses.</summary>
    public IReadOnlyList<string> ProxyAddresses { get; }
    /// <summary>Distribution-list member distinguished names when present in the offline snapshot.</summary>
    public IReadOnlyList<string> MemberDistinguishedNames { get; }
    /// <summary>Distribution lists containing this object when present in the offline snapshot.</summary>
    public IReadOnlyList<string> MemberOfDistinguishedNames { get; }
    /// <summary>Properties truncated or omitted by OAB generation limits.</summary>
    public IReadOnlyList<uint> TruncatedPropertyTags { get; }
    /// <summary>Raw MAPI object type.</summary>
    public uint? RawObjectType { get; }
    /// <summary>Common object type projection.</summary>
    public OfflineAddressBookObjectType ObjectType { get; }
    /// <summary>MAPI display type.</summary>
    public uint? DisplayType { get; }
    /// <summary>Extended MAPI display type.</summary>
    public uint? DisplayTypeEx { get; }
    /// <summary>Expanded distribution-list member count when supplied.</summary>
    public uint? DistributionListMemberCount { get; }
    /// <summary>External distribution-list member count when supplied.</summary>
    public uint? DistributionListExternalMemberCount { get; }
    /// <summary>Whether a distribution list represents a departmental group.</summary>
    public bool? IsHierarchicalGroup { get; }
    /// <summary>Whether the target can receive rich message content.</summary>
    public bool? CanReceiveRichContent { get; }
    /// <summary>Whether the object is a distribution list.</summary>
    public bool IsDistributionList => ObjectType == OfflineAddressBookObjectType.DistributionList;

    /// <summary>Returns whether OAB generation reported a property as truncated or omitted.</summary>
    public bool IsPropertyTruncated(uint propertyTag) => TruncatedPropertyTags.Contains(propertyTag);

    /// <summary>Creates a small listing and search projection for this entry.</summary>
    public OfflineAddressBookEntrySummary ToSummary() => new OfflineAddressBookEntrySummary(this);

    /// <summary>Projects the directory identity into the shared OfficeIMO email-address model.</summary>
    public EmailAddress ToEmailAddress() {
        string? address = SmtpAddress ?? TargetAddress ?? X500Address;
        return new EmailAddress(address, DisplayName, X500Address) {
            AddressType = !string.IsNullOrWhiteSpace(SmtpAddress) ? "SMTP" :
                !string.IsNullOrWhiteSpace(X500Address) ? "EX" : null
        };
    }

    /// <summary>Projects compatible person/contact fields into the shared Outlook contact model.</summary>
    public OutlookContact ToOutlookContact() {
        var contact = new OutlookContact {
            DisplayName = DisplayName,
            GivenName = GivenName,
            Surname = Surname,
            Initials = Initials,
            CompanyName = CompanyName,
            JobTitle = JobTitle,
            Department = Department,
            OfficeLocation = OfficeLocation,
            AssistantName = AssistantName
        };
        contact.Email1.Address = SmtpAddress ?? TargetAddress ?? X500Address;
        contact.Email1.DisplayName = DisplayName;
        contact.Email1.OriginalDisplayName = DisplayName;
        contact.Email1.AddressType = !string.IsNullOrWhiteSpace(SmtpAddress) ? "SMTP" : "EX";
        contact.BusinessAddress.Street = StreetAddress;
        contact.BusinessAddress.City = Locality;
        contact.BusinessAddress.StateOrProvince = StateOrProvince;
        contact.BusinessAddress.PostalCode = PostalCode;
        contact.BusinessAddress.Country = Country;
        contact.Phones.Business = BusinessTelephone;
        contact.Phones.Business2 = BusinessTelephone2.FirstOrDefault();
        contact.Phones.Home = HomeTelephone;
        contact.Phones.Home2 = HomeTelephone2.FirstOrDefault();
        contact.Phones.Mobile = MobileTelephone;
        contact.Phones.PrimaryFax = PrimaryFax;
        contact.Phones.Assistant = AssistantTelephone;
        contact.Phones.Pager = PagerTelephone;
        return contact;
    }
}
