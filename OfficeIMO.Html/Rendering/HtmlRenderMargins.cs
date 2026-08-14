namespace OfficeIMO.Html;

/// <summary>
/// Immutable HTML render margins measured in CSS pixels.
/// </summary>
public readonly struct HtmlRenderMargins : IEquatable<HtmlRenderMargins> {
    /// <summary>Creates page or continuous-surface margins.</summary>
    public HtmlRenderMargins(double left, double top, double right, double bottom) {
        Validate(left, nameof(left));
        Validate(top, nameof(top));
        Validate(right, nameof(right));
        Validate(bottom, nameof(bottom));
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
    }

    private HtmlRenderMargins(double left, double top, double right, double bottom, bool allowNegative) {
        ValidateFinite(left, nameof(left));
        ValidateFinite(top, nameof(top));
        ValidateFinite(right, nameof(right));
        ValidateFinite(bottom, nameof(bottom));
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
    }

    /// <summary>Left margin in CSS pixels.</summary>
    public double Left { get; }

    /// <summary>Top margin in CSS pixels.</summary>
    public double Top { get; }

    /// <summary>Right margin in CSS pixels.</summary>
    public double Right { get; }

    /// <summary>Bottom margin in CSS pixels.</summary>
    public double Bottom { get; }

    /// <summary>Creates equal margins on every side.</summary>
    public static HtmlRenderMargins All(double value) => new HtmlRenderMargins(value, value, value, value);

    internal static HtmlRenderMargins FromCssPageRule(double left, double top, double right, double bottom) =>
        new HtmlRenderMargins(left, top, right, bottom, allowNegative: true);

    /// <inheritdoc />
    public bool Equals(HtmlRenderMargins other) =>
        Left.Equals(other.Left) && Top.Equals(other.Top) && Right.Equals(other.Right) && Bottom.Equals(other.Bottom);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is HtmlRenderMargins other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            hash = (hash * 31) + Left.GetHashCode();
            hash = (hash * 31) + Top.GetHashCode();
            hash = (hash * 31) + Right.GetHashCode();
            hash = (hash * 31) + Bottom.GetHashCode();
            return hash;
        }
    }

    /// <summary>Equality operator.</summary>
    public static bool operator ==(HtmlRenderMargins left, HtmlRenderMargins right) => left.Equals(right);

    /// <summary>Inequality operator.</summary>
    public static bool operator !=(HtmlRenderMargins left, HtmlRenderMargins right) => !left.Equals(right);

    private static void Validate(double value, string parameterName) {
        if (value < 0D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "HTML render margins must be finite non-negative values.");
        }
    }

    private static void ValidateFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "HTML render margins must be finite values.");
        }
    }
}
