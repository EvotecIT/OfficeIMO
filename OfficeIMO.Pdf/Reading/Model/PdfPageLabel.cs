namespace OfficeIMO.Pdf;

/// <summary>
/// A page-label rule discovered from a simple catalog page-label number tree.
/// </summary>
public sealed class PdfPageLabel {
    internal PdfPageLabel(int startPageIndex, string? style, string? prefix, int? startNumber) {
        StartPageIndex = startPageIndex;
        Style = style;
        Prefix = prefix;
        StartNumber = startNumber;
    }

    /// <summary>Zero-based page index where this label rule starts.</summary>
    public int StartPageIndex { get; }

    /// <summary>One-based page number where this label rule starts.</summary>
    public int StartPageNumber => StartPageIndex + 1;

    /// <summary>PDF label style name, for example D, R, r, A, or a, when present.</summary>
    public string? Style { get; }

    /// <summary>Optional label prefix.</summary>
    public string? Prefix { get; }

    /// <summary>Starting label number when present.</summary>
    public int? StartNumber { get; }
}
