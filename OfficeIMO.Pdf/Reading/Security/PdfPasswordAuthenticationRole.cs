namespace OfficeIMO.Pdf;

/// <summary>Identifies how the supplied password or recipient certificate authenticated with the PDF security handler.</summary>
public enum PdfPasswordAuthenticationRole {
    /// <summary>No password authentication was performed.</summary>
    None = 0,

    /// <summary>The supplied password authenticated as the document user password.</summary>
    User = 1,

    /// <summary>The supplied password authenticated as the document owner password.</summary>
    Owner = 2,

    /// <summary>
    /// A supplied certificate opened the public-key security handler as one of the document recipients.
    /// The permissions stored for that recipient apply.
    /// </summary>
    PublicKeyRecipient = 3
}
