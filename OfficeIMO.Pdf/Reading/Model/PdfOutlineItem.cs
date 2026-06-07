namespace OfficeIMO.Pdf;

/// <summary>
/// A PDF outline/bookmark entry discovered from the document catalog.
/// </summary>
public sealed class PdfOutlineItem {
    internal PdfOutlineItem(
        string title,
        int level,
        int? pageNumber,
        double? destinationTop,
        bool isExpanded,
        IReadOnlyList<PdfOutlineItem> children,
        PdfOpenActionDestinationMode? destinationMode = null,
        double? destinationLeft = null,
        double? destinationBottom = null,
        double? destinationRight = null) {
        Title = title;
        Level = level;
        PageNumber = pageNumber;
        DestinationTop = destinationTop;
        IsExpanded = isExpanded;
        Children = children;
        DestinationMode = destinationMode;
        DestinationLeft = destinationLeft;
        DestinationBottom = destinationBottom;
        DestinationRight = destinationRight;
    }

    /// <summary>Bookmark text shown by PDF readers.</summary>
    public string Title { get; }

    /// <summary>One-based outline nesting level.</summary>
    public int Level { get; }

    /// <summary>One-based target page number when the destination could be resolved.</summary>
    public int? PageNumber { get; }

    /// <summary>Top coordinate of the destination when present.</summary>
    public double? DestinationTop { get; }

    /// <summary>Left coordinate of the destination when present.</summary>
    public double? DestinationLeft { get; }

    /// <summary>Bottom coordinate of the destination rectangle when present.</summary>
    public double? DestinationBottom { get; }

    /// <summary>Right coordinate of the destination rectangle when present.</summary>
    public double? DestinationRight { get; }

    /// <summary>Viewer destination mode when the outline uses a supported destination array.</summary>
    public PdfOpenActionDestinationMode? DestinationMode { get; }

    /// <summary>Whether this outline entry is expanded according to its PDF count state.</summary>
    public bool IsExpanded { get; }

    /// <summary>Nested outline entries.</summary>
    public IReadOnlyList<PdfOutlineItem> Children { get; }
}
