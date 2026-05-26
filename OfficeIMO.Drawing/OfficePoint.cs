using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free two-dimensional point used by shared drawing descriptors.
/// Coordinates are expressed in the caller's layout unit.
/// </summary>
public struct OfficePoint : IEquatable<OfficePoint> {
    /// <summary>Horizontal coordinate.</summary>
    public double X { get; }

    /// <summary>Vertical coordinate.</summary>
    public double Y { get; }

    /// <summary>Creates a point.</summary>
    public OfficePoint(double x, double y) {
        X = x;
        Y = y;
    }

    /// <inheritdoc />
    public bool Equals(OfficePoint other) => X.Equals(other.X) && Y.Equals(other.Y);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficePoint other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            return (X.GetHashCode() * 397) ^ Y.GetHashCode();
        }
    }

    /// <summary>Compares two points for equality.</summary>
    public static bool operator ==(OfficePoint left, OfficePoint right) => left.Equals(right);

    /// <summary>Compares two points for inequality.</summary>
    public static bool operator !=(OfficePoint left, OfficePoint right) => !left.Equals(right);
}
