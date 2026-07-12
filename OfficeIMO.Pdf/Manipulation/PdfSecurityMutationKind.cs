namespace OfficeIMO.Pdf;

/// <summary>Existing-document Standard security mutation performed by <see cref="PdfSecurityEditor"/>.</summary>
public enum PdfSecurityMutationKind {
    /// <summary>Standard password security was added to an unencrypted document.</summary>
    Encrypt = 0,

    /// <summary>Standard password security was removed with owner authorization.</summary>
    Decrypt = 1,

    /// <summary>Standard password security settings were replaced with owner authorization.</summary>
    Reencrypt = 2
}
