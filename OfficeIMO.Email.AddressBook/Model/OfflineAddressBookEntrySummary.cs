namespace OfficeIMO.Email.AddressBook;

/// <summary>Small projection suitable for listings and bounded search results.</summary>
public sealed class OfflineAddressBookEntrySummary {
    internal OfflineAddressBookEntrySummary(OfflineAddressBookEntry entry) {
        Reference = entry.Reference;
        DisplayName = entry.DisplayName;
        SmtpAddress = entry.SmtpAddress;
        Account = entry.Account;
        CompanyName = entry.CompanyName;
        Department = entry.Department;
        ObjectType = entry.ObjectType;
        IsDistributionList = entry.IsDistributionList;
    }

    /// <summary>Stable reference for an explicit full read.</summary>
    public OfflineAddressBookEntryReference Reference { get; }
    /// <summary>Display name.</summary>
    public string? DisplayName { get; }
    /// <summary>Primary SMTP address.</summary>
    public string? SmtpAddress { get; }
    /// <summary>Directory account.</summary>
    public string? Account { get; }
    /// <summary>Company name.</summary>
    public string? CompanyName { get; }
    /// <summary>Department.</summary>
    public string? Department { get; }
    /// <summary>Projected object type.</summary>
    public OfflineAddressBookObjectType ObjectType { get; }
    /// <summary>Whether this entry is a distribution list.</summary>
    public bool IsDistributionList { get; }
}
