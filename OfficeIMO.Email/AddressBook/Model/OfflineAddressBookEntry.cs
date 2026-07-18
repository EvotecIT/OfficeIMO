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

        X500Address = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.EmailAddress);
        SmtpAddress = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.SmtpAddress);
        DisplayName = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.DisplayName) ??
            OabPropertyValues.String(properties, MapiKnownProperties.PidTag.DisplayNamePrintable);
        Account = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.Account);
        GivenName = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.GivenName);
        Surname = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.Surname);
        Initials = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.Initials);
        CompanyName = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.CompanyName);
        JobTitle = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.Title);
        Department = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.DepartmentName);
        OfficeLocation = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.OfficeLocation);
        AssistantName = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.Assistant);
        Comment = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.Comment);
        TargetAddress = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.AddressBookTargetAddress);
        StreetAddress = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.StreetAddress);
        Locality = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.Locality);
        StateOrProvince = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.StateOrProvince);
        PostalCode = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.PostalCode);
        Country = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.Country);
        BusinessTelephone = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.BusinessTelephoneNumber);
        HomeTelephone = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.HomeTelephoneNumber);
        MobileTelephone = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.MobileTelephoneNumber);
        PrimaryFax = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.PrimaryFaxNumber);
        AssistantTelephone = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.AssistantTelephoneNumber);
        PagerTelephone = OabPropertyValues.String(properties, MapiKnownProperties.PidTag.PagerTelephoneNumber);
        BusinessTelephone2 = OabPropertyValues.Strings(properties, MapiKnownProperties.PidTag.Business2TelephoneNumbers);
        HomeTelephone2 = OabPropertyValues.Strings(properties, MapiKnownProperties.PidTag.Home2TelephoneNumbers);
        ProxyAddresses = OabPropertyValues.Strings(properties, MapiKnownProperties.PidTag.AddressBookProxyAddresses);
        MemberDistinguishedNames = OabPropertyValues.Strings(properties, MapiKnownProperties.PidTag.AddressBookMembers);
        MemberOfDistinguishedNames = OabPropertyValues.Strings(properties, MapiKnownProperties.PidTag.AddressBookMemberOf);
        TruncatedPropertyTags = OabPropertyValues.UInt32s(properties,
            MapiKnownProperties.PidTag.OfflineAddressBookTruncatedProperties);
        RawObjectType = OabPropertyValues.UInt32(properties, MapiKnownProperties.PidTag.ObjectType);
        DisplayType = OabPropertyValues.UInt32(properties, MapiKnownProperties.PidTag.DisplayType);
        DisplayTypeEx = OabPropertyValues.UInt32(properties, MapiKnownProperties.PidTag.DisplayTypeEx);
        DistributionListMemberCount = OabPropertyValues.UInt32(properties,
            MapiKnownProperties.PidTag.AddressBookDistributionListMemberCount);
        DistributionListExternalMemberCount = OabPropertyValues.UInt32(properties,
            MapiKnownProperties.PidTag.AddressBookDistributionListExternalMemberCount);
        IsHierarchicalGroup = OabPropertyValues.Boolean(properties,
            MapiKnownProperties.PidTag.AddressBookHierarchicalGroup);
        CanReceiveRichContent = OabPropertyValues.Boolean(properties, MapiKnownProperties.PidTag.SendRichInfo);
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
