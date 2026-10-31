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
    Bottom,
    /// <summary>Top on odd output pages and bottom on even output pages.</summary>
    Inside,
    /// <summary>Bottom on odd output pages and top on even output pages.</summary>
    Outside
}

/// <summary>
/// Immutable placement and text clearance for a floating table. Distances are physical lengths;
/// positive vertical offsets move down the page. Following paragraphs wrap around the table.
/// </summary>
/// <remarks>Deferred tables support top alignment because their total height is not known before streaming their rows.</remarks>
public sealed class PdfTablePosition {
    /// <summary>Creates placement relative to the selected anchor rectangles.</summary>
    public PdfTablePosition(PdfTableAnchor horizontalAnchor = PdfTableAnchor.Margin,
        PdfTableAnchor verticalAnchor = PdfTableAnchor.Flow, PdfAlign horizontalAlignment = PdfAlign.Left,
        PdfTableVerticalAlignment verticalAlignment = PdfTableVerticalAlignment.Top,
        global::OfficeIMO.OfficeLength horizontalOffset = default, global::OfficeIMO.OfficeLength verticalOffset = default,
        global::OfficeIMO.OfficeLength distanceLeft = default, global::OfficeIMO.OfficeLength distanceRight = default, global::OfficeIMO.OfficeLength distanceTop = default, global::OfficeIMO.OfficeLength distanceBottom = default,
        bool mirrorHorizontalOnEvenPages = false, bool allowOverlap = true) {
        if (horizontalAnchor < PdfTableAnchor.Flow || horizontalAnchor > PdfTableAnchor.Page) throw new ArgumentOutOfRangeException(nameof(horizontalAnchor));
        if (verticalAnchor < PdfTableAnchor.Flow || verticalAnchor > PdfTableAnchor.Page) throw new ArgumentOutOfRangeException(nameof(verticalAnchor));
        if (horizontalAlignment != PdfAlign.Left && horizontalAlignment != PdfAlign.Center && horizontalAlignment != PdfAlign.Right) throw new ArgumentOutOfRangeException(nameof(horizontalAlignment));
        if (verticalAlignment < PdfTableVerticalAlignment.Top || verticalAlignment > PdfTableVerticalAlignment.Outside) throw new ArgumentOutOfRangeException(nameof(verticalAlignment));
        Validate(horizontalOffset.ToPoints(), nameof(horizontalOffset), false);
        Validate(verticalOffset.ToPoints(), nameof(verticalOffset), false);
        Validate(distanceLeft.ToPoints(), nameof(distanceLeft), true);
        Validate(distanceRight.ToPoints(), nameof(distanceRight), true);
        Validate(distanceTop.ToPoints(), nameof(distanceTop), true);
        Validate(distanceBottom.ToPoints(), nameof(distanceBottom), true);
        MirrorHorizontalOnEvenPages = mirrorHorizontalOnEvenPages; AllowOverlap = allowOverlap;
        HorizontalAnchor = horizontalAnchor; VerticalAnchor = verticalAnchor;
        HorizontalAlignment = horizontalAlignment; VerticalAlignment = verticalAlignment;
        HorizontalOffset = horizontalOffset; VerticalOffset = verticalOffset;
        DistanceLeft = distanceLeft; DistanceRight = distanceRight;
        DistanceTop = distanceTop; DistanceBottom = distanceBottom;
    }

    /// <summary>Reverse left/right alignment on even output pages for inside/outside placement.</summary>
    public bool MirrorHorizontalOnEvenPages { get; }
    /// <summary>Whether this table can intersect earlier floating table bounds.</summary>
    public bool AllowOverlap { get; }

    /// <summary>Horizontal reference rectangle.</summary>
    public PdfTableAnchor HorizontalAnchor { get; }
    /// <summary>Vertical reference rectangle.</summary>
    public PdfTableAnchor VerticalAnchor { get; }
    /// <summary>Horizontal alignment in the reference rectangle.</summary>
    public PdfAlign HorizontalAlignment { get; }
    /// <summary>Vertical alignment in the reference rectangle.</summary>
    public PdfTableVerticalAlignment VerticalAlignment { get; }
    /// <summary>Horizontal offset to the right, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength HorizontalOffset { get; }
    /// <summary>Vertical offset downwards, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength VerticalOffset { get; }
    /// <summary>Text clearance at the left edge, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength DistanceLeft { get; }
    /// <summary>Text clearance at the right edge, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength DistanceRight { get; }
    /// <summary>Text clearance at the top edge, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength DistanceTop { get; }
    /// <summary>Text clearance at the bottom edge, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength DistanceBottom { get; }

    private static void Validate(double value, string name, bool nonNegative) {
        if (double.IsNaN(value) || double.IsInfinity(value) || nonNegative && value < 0)
            throw new ArgumentOutOfRangeException(name);
    }
}
