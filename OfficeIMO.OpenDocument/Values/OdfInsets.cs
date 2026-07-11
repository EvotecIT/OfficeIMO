namespace OfficeIMO.OpenDocument;

/// <summary>Top, right, bottom, and left ODF lengths.</summary>
public readonly struct OdfInsets : IEquatable<OdfInsets> {
    /// <summary>Creates four-sided insets.</summary>
    public OdfInsets(OdfLength top, OdfLength right, OdfLength bottom, OdfLength left) {
        Top = top; Right = right; Bottom = bottom; Left = left;
    }
    /// <summary>Top inset.</summary>
    public OdfLength Top { get; }
    /// <summary>Right inset.</summary>
    public OdfLength Right { get; }
    /// <summary>Bottom inset.</summary>
    public OdfLength Bottom { get; }
    /// <summary>Left inset.</summary>
    public OdfLength Left { get; }
    /// <summary>Creates centimeter insets.</summary>
    public static OdfInsets FromCentimeters(double top, double right, double bottom, double left) => new OdfInsets(
        OdfLength.Centimeters(top), OdfLength.Centimeters(right), OdfLength.Centimeters(bottom), OdfLength.Centimeters(left));
    /// <inheritdoc />
    public bool Equals(OdfInsets other) => Top.Equals(other.Top) && Right.Equals(other.Right) && Bottom.Equals(other.Bottom) && Left.Equals(other.Left);
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OdfInsets other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => Top.GetHashCode() ^ Right.GetHashCode() ^ Bottom.GetHashCode() ^ Left.GetHashCode();
}
