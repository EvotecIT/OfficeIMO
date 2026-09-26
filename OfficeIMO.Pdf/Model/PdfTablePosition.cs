namespace OfficeIMO.Pdf;

/// <summary>Reference rectangle used to anchor a floating table.</summary>
public enum PdfTableAnchor {
    /// <summary>The current text frame and vertical flow cursor.</summary>
    Flow,
    /// <summary>The page content rectangle, excluding margins.</summary>
    Margin,
    /// <summary>The entire page rectangle.</summary>
    Page
}

/// <summary>Vertical alignment within the table's anchor rectangle.</summary>
public enum PdfTableVerticalAlignment {
    /// <summary>Align the table's top edge.</summary>
    Top,
    /// <summary>Center the table vertically.</summary>
    Center,
    /// <summary>Align the table's bottom edge.</summary>
    Bottom
}

/// <summary>
/// Immutable placement and text clearance for a floating table. Distances are in points;
/// positive vertical offsets move down the page. Following paragraphs wrap around the table.
/// </summary>
/// <remarks>Deferred tables support top alignment because their total height is not known before streaming their rows.</remarks>
public sealed class PdfTablePosition {
    /// <summary>Creates placement relative to the selected anchor rectangles.</summary>
    public PdfTablePosition(PdfTableAnchor horizontalAnchor = PdfTableAnchor.Margin,
        PdfTableAnchor verticalAnchor = PdfTableAnchor.Flow, PdfAlign horizontalAlignment = PdfAlign.Left,
        PdfTableVerticalAlignment verticalAlignment = PdfTableVerticalAlignment.Top,
        double horizontalOffset = 0, double verticalOffset = 0,
        double distanceLeft = 0, double distanceRight = 0, double distanceTop = 0, double distanceBottom = 0) {
        if (horizontalAnchor < PdfTableAnchor.Flow || horizontalAnchor > PdfTableAnchor.Page) throw new ArgumentOutOfRangeException(nameof(horizontalAnchor));
        if (verticalAnchor < PdfTableAnchor.Flow || verticalAnchor > PdfTableAnchor.Page) throw new ArgumentOutOfRangeException(nameof(verticalAnchor));
        if (horizontalAlignment != PdfAlign.Left && horizontalAlignment != PdfAlign.Center && horizontalAlignment != PdfAlign.Right) throw new ArgumentOutOfRangeException(nameof(horizontalAlignment));
        if (verticalAlignment < PdfTableVerticalAlignment.Top || verticalAlignment > PdfTableVerticalAlignment.Bottom) throw new ArgumentOutOfRangeException(nameof(verticalAlignment));
        Validate(horizontalOffset, nameof(horizontalOffset), false);
        Validate(verticalOffset, nameof(verticalOffset), false);
        Validate(distanceLeft, nameof(distanceLeft), true);
        Validate(distanceRight, nameof(distanceRight), true);
        Validate(distanceTop, nameof(distanceTop), true);
        Validate(distanceBottom, nameof(distanceBottom), true);
        HorizontalAnchor = horizontalAnchor; VerticalAnchor = verticalAnchor;
        HorizontalAlignment = horizontalAlignment; VerticalAlignment = verticalAlignment;
        HorizontalOffset = horizontalOffset; VerticalOffset = verticalOffset;
        DistanceLeft = distanceLeft; DistanceRight = distanceRight;
        DistanceTop = distanceTop; DistanceBottom = distanceBottom;
    }

    /// <summary>Horizontal reference rectangle.</summary>
    public PdfTableAnchor HorizontalAnchor { get; }
    /// <summary>Vertical reference rectangle.</summary>
    public PdfTableAnchor VerticalAnchor { get; }
    /// <summary>Horizontal alignment in the reference rectangle.</summary>
    public PdfAlign HorizontalAlignment { get; }
    /// <summary>Vertical alignment in the reference rectangle.</summary>
    public PdfTableVerticalAlignment VerticalAlignment { get; }
    /// <summary>Horizontal offset to the right, in points.</summary>
    public double HorizontalOffset { get; }
    /// <summary>Vertical offset downwards, in points.</summary>
    public double VerticalOffset { get; }
    /// <summary>Text clearance at the left edge, in points.</summary>
    public double DistanceLeft { get; }
    /// <summary>Text clearance at the right edge, in points.</summary>
    public double DistanceRight { get; }
    /// <summary>Text clearance at the top edge, in points.</summary>
    public double DistanceTop { get; }
    /// <summary>Text clearance at the bottom edge, in points.</summary>
    public double DistanceBottom { get; }

    private static void Validate(double value, string name, bool nonNegative) {
        if (double.IsNaN(value) || double.IsInfinity(value) || nonNegative && value < 0)
            throw new ArgumentOutOfRangeException(name);
    }
}
