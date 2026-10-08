namespace OfficeIMO.Pdf;

/// <summary>Controls spacing and wrapping of ordinary spaces in generated flow text.</summary>
public enum PdfTextWhitespaceMode {
    /// <summary>Collapses repeated spaces and discards spaces at the start of a line.</summary>
    Collapse,
    /// <summary>Preserves spaces within a line and discards excess spaces when they wrap.</summary>
    Preserve,
    /// <summary>Preserves literal spacing, including additional blank lines when spaces wrap.</summary>
    Preformatted
}
