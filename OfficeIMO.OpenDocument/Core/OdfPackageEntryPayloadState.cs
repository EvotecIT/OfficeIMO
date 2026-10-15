namespace OfficeIMO.OpenDocument;

/// <summary>Native package payload materialization state. It does not establish cryptographic authenticity.</summary>
public enum OdfPackageEntryPayloadState {
    /// <summary>Declared ZIP size and path were checked; expansion and complete decoded-length checks await explicit access.</summary>
    Deferred,
    /// <summary>Complete bytes passed the ordinary native entry read checks.</summary>
    Loaded,
    /// <summary>Complete bytes passed the supported ODF decryption, checksum and declared-length checks.</summary>
    Decrypted
}
