namespace OfficeIMO.Email.AddressBook;

/// <summary>Depth of an explicit OAB validation pass.</summary>
public enum OfflineAddressBookValidationMode {
    /// <summary>Validate the file checksum without walking entry records.</summary>
    ChecksumOnly,
    /// <summary>Validate record sizes and boundaries without decoding property values.</summary>
    Framing,
    /// <summary>Validate framing and decode every selected record against the file-defined schema.</summary>
    FullDecode
}
