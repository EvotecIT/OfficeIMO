namespace OfficeIMO.Pdf;

/// <summary>
/// A PDF outline/bookmark entry discovered from the document catalog.
/// </summary>
public sealed class PdfOutlineItem {
    internal PdfOutlineItem(string title, int level, int? pageNumber, double? destinationTop, IReadOnlyList<PdfOutlineItem> children) {
        Title = title;
        Level = level;
        PageNumber = pageNumber;
        DestinationTop = destinationTop;
        Children = children;
    }

    /// <summary>Bookmark text shown by PDF readers.</summary>
    public string Title { get; }

    /// <summary>One-based outline nesting level.</summary>
    public int Level { get; }

    /// <summary>One-based target page number when the destination could be resolved.</summary>
    public int? PageNumber { get; }

    /// <summary>Top coordinate of the destination when present.</summary>
    public double? DestinationTop { get; }

    /// <summary>Nested outline entries.</summary>
    public IReadOnlyList<PdfOutlineItem> Children { get; }
}
