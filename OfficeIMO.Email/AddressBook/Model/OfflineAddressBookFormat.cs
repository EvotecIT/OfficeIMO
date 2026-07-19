namespace OfficeIMO.Email.AddressBook;

/// <summary>Recognized role of an Outlook Offline Address Book component.</summary>
public enum OfflineAddressBookFormat {
    /// <summary>Unrecognized .oab component.</summary>
    Unknown = 0,
    /// <summary>Uncompressed OAB version 4 Full Details file.</summary>
    Version4FullDetails = 1,
    /// <summary>OAB display template.</summary>
    DisplayTemplate = 2,
    /// <summary>Legacy version 2 or version 3 Browse file.</summary>
    LegacyBrowse = 3,
    /// <summary>Legacy version 2 or version 3 ambiguous-name-resolution index.</summary>
    LegacyAnrIndex = 4,
    /// <summary>Legacy version 2 or version 3 relative-distinguished-name index.</summary>
    LegacyRdnIndex = 5,
    /// <summary>Legacy version 2 or version 3 Details file.</summary>
    LegacyDetails = 6,
    /// <summary>Legacy version 2 or version 3 Changes file.</summary>
    LegacyChanges = 7
}
