using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Immutable point in an HTML render surface's CSS-pixel coordinate space.</summary>
public readonly struct HtmlRenderPoint : IEquatable<HtmlRenderPoint> {
    /// <summary>Creates a finite point.</summary>
    public HtmlRenderPoint(double x, double y) {
        ValidateFinite(x, nameof(x));
        ValidateFinite(y, nameof(y));
        X = x;
        Y = y;
    }

    /// <summary>Horizontal coordinate in CSS pixels.</summary>
    public double X { get; }
    /// <summary>Vertical coordinate in CSS pixels.</summary>
    public double Y { get; }

    /// <inheritdoc />
    public bool Equals(HtmlRenderPoint other) => X.Equals(other.X) && Y.Equals(other.Y);
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is HtmlRenderPoint other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() { unchecked { return (X.GetHashCode() * 397) ^ Y.GetHashCode(); } }
    /// <summary>Compares two points.</summary>
    public static bool operator ==(HtmlRenderPoint left, HtmlRenderPoint right) => left.Equals(right);
    /// <summary>Compares two points.</summary>
    public static bool operator !=(HtmlRenderPoint left, HtmlRenderPoint right) => !left.Equals(right);

    internal OfficePoint ToOfficePoint() => new OfficePoint(X, Y);

    private static void ValidateFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Render coordinates must be finite.");
        }
    }
}

/// <summary>Immutable axis-aligned rectangle in CSS-pixel coordinates.</summary>
public readonly struct HtmlRenderRectangle : IEquatable<HtmlRenderRectangle> {
    /// <summary>Creates a finite rectangle with non-negative dimensions.</summary>
    public HtmlRenderRectangle(double x, double y, double width, double height) {
        ValidateFinite(x, nameof(x));
        ValidateFinite(y, nameof(y));
        ValidateDimension(width, nameof(width));
        ValidateDimension(height, nameof(height));
        X = x;
        Y = y;
        Width = width;
        Height = height;
    }

    /// <summary>Left coordinate.</summary>
    public double X { get; }
    /// <summary>Top coordinate.</summary>
    public double Y { get; }
    /// <summary>Width.</summary>
    public double Width { get; }
    /// <summary>Height.</summary>
    public double Height { get; }
    /// <summary>Right coordinate.</summary>
    public double Right => X + Width;
    /// <summary>Bottom coordinate.</summary>
    public double Bottom => Y + Height;

    /// <summary>Returns whether the point lies in the closed rectangle.</summary>
    public bool Contains(HtmlRenderPoint point) =>
        point.X >= X && point.X <= Right && point.Y >= Y && point.Y <= Bottom;

    /// <summary>Returns whether this rectangle overlaps another rectangle.</summary>
    public bool Intersects(HtmlRenderRectangle other) =>
        Right >= other.X && other.Right >= X && Bottom >= other.Y && other.Bottom >= Y;

    /// <inheritdoc />
    public bool Equals(HtmlRenderRectangle other) =>
        X.Equals(other.X) && Y.Equals(other.Y) && Width.Equals(other.Width) && Height.Equals(other.Height);
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is HtmlRenderRectangle other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = X.GetHashCode();
            hash = (hash * 397) ^ Y.GetHashCode();
            hash = (hash * 397) ^ Width.GetHashCode();
            hash = (hash * 397) ^ Height.GetHashCode();
            return hash;
        }
    }
    /// <summary>Compares two rectangles.</summary>
    public static bool operator ==(HtmlRenderRectangle left, HtmlRenderRectangle right) => left.Equals(right);
    /// <summary>Compares two rectangles.</summary>
    public static bool operator !=(HtmlRenderRectangle left, HtmlRenderRectangle right) => !left.Equals(right);

    internal static HtmlRenderRectangle Transform(OfficeTransform transform, HtmlRenderRectangle rectangle) {
        var bounds = transform.TransformRectangleBounds(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
        return new HtmlRenderRectangle(bounds.Left, bounds.Top, bounds.Right - bounds.Left, bounds.Bottom - bounds.Top);
    }

    private static void ValidateFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Render coordinates must be finite.");
        }
    }

    private static void ValidateDimension(double value, string parameterName) {
        if (value < 0D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Render dimensions must be finite non-negative numbers.");
        }
    }
}

/// <summary>Maps one output-surface point back to a source page or source-canvas slice.</summary>
public sealed class HtmlRenderSourcePoint {
    internal HtmlRenderSourcePoint(int outputIndex, int sourcePageNumber, HtmlRenderPoint outputPoint,
        HtmlRenderPoint sourcePoint, bool clipped) {
        OutputIndex = outputIndex;
        SourcePageNumber = sourcePageNumber;
        OutputPoint = outputPoint;
        SourcePoint = sourcePoint;
        IsClipped = clipped;
    }

    /// <summary>Zero-based output surface index.</summary>
    public int OutputIndex { get; }
    /// <summary>One-based source page number.</summary>
    public int SourcePageNumber { get; }
    /// <summary>Point in the retained output surface.</summary>
    public HtmlRenderPoint OutputPoint { get; }
    /// <summary>Point in the original source page or completed continuous canvas.</summary>
    public HtmlRenderPoint SourcePoint { get; }
    /// <summary>Whether the contributing source placement was clipped.</summary>
    public bool IsClipped { get; }
}
