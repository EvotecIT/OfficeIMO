namespace OfficeIMO.Pdf;

/// <summary>Identifies how text wraps around a picture exclusion.</summary>
internal enum PdfImageTextWrapKind {
    Square,
    TopAndBottom,
    /// <summary>Text follows the picture's wrap polygon on each line.</summary>
    Tight,
    /// <summary>Lays out as Tight: each line clears the polygon's extent across that line.</summary>
    Through
}

/// <summary>Retains the producer's requested side of a rectangular picture exclusion.</summary>
internal enum PdfImageTextWrapSide {
    BothSides,
    Left,
    Right,
    Largest
}

/// <summary>
/// Source text distances in points. These expand the exclusion, never the picture's paint frame.
/// Side is ignored by TopAndBottom because that mode clears the full paragraph width.
/// Tight and Through carry a closed wrap polygon whose points are fractions of the picture's
/// width and height, measured from its top-left corner; points may lie outside the picture.
/// </summary>
internal sealed class PdfImageTextExclusion {
    internal PdfImageTextExclusion(PdfImageTextWrapKind kind, PdfImageTextWrapSide side,
        double distanceLeft, double distanceTop, double distanceRight, double distanceBottom,
        IReadOnlyList<(double X, double Y)>? polygon = null) {
        Guard.NonNegative(distanceLeft, nameof(distanceLeft));
        Guard.NonNegative(distanceTop, nameof(distanceTop));
        Guard.NonNegative(distanceRight, nameof(distanceRight));
        Guard.NonNegative(distanceBottom, nameof(distanceBottom));
        bool polygonKind = kind is PdfImageTextWrapKind.Tight or PdfImageTextWrapKind.Through;
        if (polygonKind != (polygon != null))
            throw new ArgumentException("Tight and through wrapping need a wrap polygon; other kinds take none.", nameof(polygon));
        if (polygon != null && (polygon.Count < 3 || polygon.Any(point => double.IsNaN(point.X) || double.IsInfinity(point.X) ||
                double.IsNaN(point.Y) || double.IsInfinity(point.Y))))
            throw new ArgumentException("A wrap polygon needs at least three finite points.", nameof(polygon));
        Polygon = polygon?.ToArray();
        Kind = kind;
        Side = side;
        DistanceLeft = distanceLeft;
        DistanceTop = distanceTop;
        DistanceRight = distanceRight;
        DistanceBottom = distanceBottom;
    }

    internal PdfImageTextWrapKind Kind { get; }
    internal PdfImageTextWrapSide Side { get; }
    internal double DistanceLeft { get; }
    internal double DistanceTop { get; }
    internal double DistanceRight { get; }
    internal double DistanceBottom { get; }
    internal IReadOnlyList<(double X, double Y)>? Polygon { get; }
}

/// <summary>
/// Source top-left placement in points. MarginX retains the authored left margin for physical-page
/// margin resolution; ParagraphY defers Y to the admitted paragraph origin before its own clearance.
/// </summary>
internal readonly record struct PdfAnchoredImagePosition(
    double X, double Y, bool MarginX, double AuthoredLeftMargin, bool ParagraphY);

/// <summary>
/// Carries one picture's local canvas paint and optional exclusion through shared layout.
/// This descriptor neither reserves flow height nor paints itself. A layout plan resolves its source
/// position once, then uses that same placement for its exclusion and exactly-once image painting.
/// </summary>
internal sealed class PdfCanvasAnchoredImageItem : PdfCanvasItem {
    internal PdfCanvasAnchoredImageItem(object sourceIdentity, PdfCanvasBlock content,
        double width, double height, PdfAnchoredImagePosition position,
        PdfImageTextExclusion? exclusion) : base(position.X, position.Y) {
        Guard.NotNull(sourceIdentity, nameof(sourceIdentity));
        Guard.NotNull(content, nameof(content));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        if (double.IsNaN(position.X) || double.IsInfinity(position.X) ||
            double.IsNaN(position.Y) || double.IsInfinity(position.Y) ||
            double.IsNaN(position.AuthoredLeftMargin) || double.IsInfinity(position.AuthoredLeftMargin))
            throw new ArgumentOutOfRangeException(nameof(position), "Anchor coordinates must be finite.");
        SourceIdentity = sourceIdentity;
        Content = content;
        Width = width;
        Height = height;
        Position = position;
        Exclusion = exclusion;
    }

    /// <summary>
    /// Opaque source identity stable across layout passes. The caller supplies original source
    /// ownership, rather than a freshly materialized callback item or a generated per-pass identifier.
    /// Page, story and occurrence scope are supplied by the layout plan.
    /// </summary>
    internal object SourceIdentity { get; }
    internal PdfCanvasBlock Content { get; }
    internal double Width { get; }
    internal double Height { get; }
    internal PdfAnchoredImagePosition Position { get; }
    internal PdfImageTextExclusion? Exclusion { get; }
}
