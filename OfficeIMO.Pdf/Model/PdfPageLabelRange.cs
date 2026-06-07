namespace OfficeIMO.Pdf;

/// <summary>
/// Defines a generated PDF page-label rule beginning at a one-based document page number.
/// </summary>
public sealed class PdfPageLabelRange {
    /// <summary>Creates a generated page-label rule.</summary>
    public PdfPageLabelRange(int startPageNumber, PdfPageNumberStyle style, int startNumber = 1, string? prefix = null) {
        if (startPageNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(startPageNumber), "PDF page-label range start page must be positive.");
        }

        Guard.PageNumberStyle(style, nameof(style));
        if (startNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(startNumber), "PDF page-label range start number must be positive.");
        }

        PdfPageLabelDictionaryBuilder.ValidatePrefix(prefix, nameof(prefix));
        StartPageNumber = startPageNumber;
        Style = style;
        StartNumber = startNumber;
        Prefix = prefix;
    }

    /// <summary>One-based document page number where this label rule begins.</summary>
    public int StartPageNumber { get; }

    /// <summary>Numbering style used by this label rule.</summary>
    public PdfPageNumberStyle Style { get; }

    /// <summary>First visible number emitted for this label rule.</summary>
    public int StartNumber { get; }

    /// <summary>Optional label prefix, for example "A-" or "Appendix ".</summary>
    public string? Prefix { get; }
}
