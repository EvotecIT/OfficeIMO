namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentInfo {
    /// <summary>Security, signature, and revision markers read from the source PDF bytes.</summary>
    public PdfDocumentSecurityInfo Security { get; }

    /// <summary>True when the document exposes encryption, signature, permission, or incremental-update markers.</summary>
    public bool HasSecurityState =>
        Security.HasEncryption ||
        Security.HasSignatures ||
        Security.HasReadableEncryptionSettings ||
        Security.HasIncrementalUpdates ||
        Security.HasDocMDPPermissions ||
        Security.HasUsageRights;
}
