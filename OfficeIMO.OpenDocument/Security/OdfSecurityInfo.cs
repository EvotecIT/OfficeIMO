namespace OfficeIMO.OpenDocument;

/// <summary>Describes password protection associated with the current OpenDocument package.</summary>
public sealed class OdfSecurityInfo {
    internal OdfSecurityInfo(bool sourceIsEncrypted) {
        SourceIsEncrypted = sourceIsEncrypted;
    }

    /// <summary>Whether the package loaded or most recently saved by this document is encrypted.</summary>
    public bool SourceIsEncrypted { get; }

    /// <summary>Whether encrypted source content was successfully decrypted into the editable model.</summary>
    public bool IsDecrypted => SourceIsEncrypted;

    /// <summary>Encryption profile implemented for password-protected ODF packages.</summary>
    public string SupportedProfile => "AES-256-CBC / PBKDF2-HMAC-SHA1 / SHA-256 start key / SHA-256-1K checksum";
}
