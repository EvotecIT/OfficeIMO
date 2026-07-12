namespace OfficeIMO.Pdf;

/// <summary>Identifies how a supplied password authenticated with the PDF Standard security handler.</summary>
public enum PdfPasswordAuthenticationRole {
    /// <summary>No password authentication was performed.</summary>
    None = 0,

    /// <summary>The supplied password authenticated as the document user password.</summary>
    User = 1,

    /// <summary>The supplied password authenticated as the document owner password.</summary>
    Owner = 2
}
