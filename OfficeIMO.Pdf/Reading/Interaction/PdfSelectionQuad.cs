namespace OfficeIMO.Pdf;

/// <summary>A four-corner selection or hit-test region in visual top-left page coordinates.</summary>
public sealed class PdfSelectionQuad {
    internal PdfSelectionQuad(PdfSelectionPoint topLeft, PdfSelectionPoint topRight, PdfSelectionPoint bottomRight, PdfSelectionPoint bottomLeft) {
        TopLeft = topLeft;
        TopRight = topRight;
        BottomRight = bottomRight;
        BottomLeft = bottomLeft;
        Left = Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomRight.X, bottomLeft.X));
        Top = Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomRight.Y, bottomLeft.Y));
        Right = Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomRight.X, bottomLeft.X));
        Bottom = Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomRight.Y, bottomLeft.Y));
    }

    /// <summary>Top-left corner.</summary>
    public PdfSelectionPoint TopLeft { get; }

    /// <summary>Top-right corner.</summary>
    public PdfSelectionPoint TopRight { get; }

    /// <summary>Bottom-right corner.</summary>
    public PdfSelectionPoint BottomRight { get; }

    /// <summary>Bottom-left corner.</summary>
    public PdfSelectionPoint BottomLeft { get; }

    /// <summary>Smallest horizontal coordinate.</summary>
    public double Left { get; }

    /// <summary>Smallest vertical coordinate.</summary>
    public double Top { get; }

    /// <summary>Largest horizontal coordinate.</summary>
    public double Right { get; }

    /// <summary>Largest vertical coordinate.</summary>
    public double Bottom { get; }

    /// <summary>Axis-aligned bounding width.</summary>
    public double Width => Right - Left;

    /// <summary>Axis-aligned bounding height.</summary>
    public double Height => Bottom - Top;

    /// <summary>Returns true when the point lies inside this quad, with optional tolerance in points.</summary>
    public bool Contains(double x, double y, double tolerance = 0D) {
        if (tolerance < 0D || double.IsNaN(tolerance) || double.IsInfinity(tolerance)) {
            throw new ArgumentOutOfRangeException(nameof(tolerance), "Hit-test tolerance must be finite and non-negative.");
        }

        if (x < Left - tolerance || x > Right + tolerance || y < Top - tolerance || y > Bottom + tolerance) {
            return false;
        }

        PdfSelectionPoint[] points = { TopLeft, TopRight, BottomRight, BottomLeft };
        bool? sign = null;
        for (int i = 0; i < points.Length; i++) {
            PdfSelectionPoint a = points[i];
            PdfSelectionPoint b = points[(i + 1) % points.Length];
            double cross = (b.X - a.X) * (y - a.Y) - (b.Y - a.Y) * (x - a.X);
            if (Math.Abs(cross) <= tolerance) {
                continue;
            }

            bool current = cross > 0D;
            if (sign.HasValue && sign.Value != current) {
                return false;
            }

            sign = current;
        }

        return true;
    }

    internal bool Intersects(double left, double top, double right, double bottom) {
        return Right >= left && Left <= right && Bottom >= top && Top <= bottom;
    }
}
