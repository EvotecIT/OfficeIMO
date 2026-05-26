using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// A color stop inside a reusable gradient descriptor.
/// </summary>
public readonly struct OfficeGradientStop : IEquatable<OfficeGradientStop> {
    /// <summary>Stop offset from 0.0 to 1.0.</summary>
    public double Offset { get; }

    /// <summary>Stop color.</summary>
    public OfficeColor Color { get; }

    /// <summary>Creates a gradient stop.</summary>
    public OfficeGradientStop(double offset, OfficeColor color) {
        if (double.IsNaN(offset) || double.IsInfinity(offset) || offset < 0D || offset > 1D) {
            throw new ArgumentOutOfRangeException(nameof(offset), "Gradient stop offsets must be finite values between 0 and 1.");
        }

        Offset = offset;
        Color = color;
    }

    /// <inheritdoc />
    public bool Equals(OfficeGradientStop other) => Offset.Equals(other.Offset) && Color.Equals(other.Color);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeGradientStop other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            return (Offset.GetHashCode() * 397) ^ Color.GetHashCode();
        }
    }

    /// <summary>Returns true when two gradient stops are equal.</summary>
    public static bool operator ==(OfficeGradientStop left, OfficeGradientStop right) => left.Equals(right);

    /// <summary>Returns true when two gradient stops are not equal.</summary>
    public static bool operator !=(OfficeGradientStop left, OfficeGradientStop right) => !left.Equals(right);
}
