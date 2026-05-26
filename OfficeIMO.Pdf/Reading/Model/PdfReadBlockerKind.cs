namespace OfficeIMO.Pdf;

/// <summary>
/// Categories of PDF inputs that OfficeIMO.Pdf cannot read yet.
/// </summary>
public enum PdfReadBlockerKind {
    /// <summary>PDF header was not found.</summary>
    MissingHeader,

    /// <summary>Encrypted PDFs cannot be read yet.</summary>
    Encryption,

    /// <summary>No page tree entries were discovered.</summary>
    NoPages,

    /// <summary>The parser could not inspect this PDF shape yet.</summary>
    ParserUnsupported
}
