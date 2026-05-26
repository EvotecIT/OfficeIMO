namespace OfficeIMO.Pdf;

/// <summary>
/// Baseline placement for rich paragraph text runs.
/// </summary>
public enum PdfTextBaseline {
    /// <summary>Use the paragraph baseline.</summary>
    Normal,
    /// <summary>Raise and scale the run like Word superscript text.</summary>
    Superscript,
    /// <summary>Lower and scale the run like Word subscript text.</summary>
    Subscript
}
