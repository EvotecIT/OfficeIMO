namespace OfficeIMO.Pdf;

/// <summary>
/// Represents an outline/bookmark-derived page range suitable for PDF splitting.
/// </summary>
public sealed class PdfBookmarkPageRange {
    internal PdfBookmarkPageRange(string title, int level, PdfPageRange pageRange, PdfOutlineItem outline) {
        Title = title;
        Level = level;
        PageRange = pageRange;
        Outline = outline;
    }

    /// <summary>Bookmark text shown by PDF readers.</summary>
    public string Title { get; }

    /// <summary>One-based outline nesting level.</summary>
    public int Level { get; }

    /// <summary>Inclusive one-based page range covered by this bookmark.</summary>
    public PdfPageRange PageRange { get; }

    /// <summary>Original outline item used to create the range.</summary>
    public PdfOutlineItem Outline { get; }
}
