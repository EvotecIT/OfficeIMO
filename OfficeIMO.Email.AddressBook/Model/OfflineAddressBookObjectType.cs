namespace OfficeIMO.Email.AddressBook;

/// <summary>Common MAPI object types stored in OAB entries.</summary>
public enum OfflineAddressBookObjectType {
    /// <summary>Unknown or provider-specific object type.</summary>
    Unknown = 0,
    /// <summary>Address-book container.</summary>
    Container = 3,
    /// <summary>Mail-enabled user or contact.</summary>
    MailUser = 6,
    /// <summary>Distribution list or group.</summary>
    DistributionList = 8
}
