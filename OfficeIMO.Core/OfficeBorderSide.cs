using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes one visible edge of a rectangular border box.
/// </summary>
public readonly struct OfficeBorderSide : IEquatable<OfficeBorderSide> {
    /// <summary>
    /// Creates a border side.
    /// </summary>
    public OfficeBorderSide(OfficeColor color, double width, OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid, OfficeBorderLineKind lineKind = OfficeBorderLineKind.Single, double doubleLineSeparation = 0D) {
        if (double.IsNaN(width) || double.IsInfinity(width) || width < 0D) {
            throw new ArgumentOutOfRangeException(nameof(width), "Border width must be a finite non-negative number.");
        }

        if (double.IsNaN(doubleLineSeparation) || double.IsInfinity(doubleLineSeparation) || doubleLineSeparation < 0D) {
            throw new ArgumentOutOfRangeException(nameof(doubleLineSeparation), "Double-line separation must be a finite non-negative number.");
        }

        Color = color;
        Width = width;
        DashStyle = dashStyle;
        LineKind = lineKind;
        DoubleLineSeparation = doubleLineSeparation;
    }

    /// <summary>Border color.</summary>
    public OfficeColor Color { get; }

    /// <summary>Border width in the caller's layout unit.</summary>
    public double Width { get; }

    /// <summary>Border dash style.</summary>
    public OfficeStrokeDashStyle DashStyle { get; }

    /// <summary>Border line composition.</summary>
    public OfficeBorderLineKind LineKind { get; }

    /// <summary>Distance between the two stroke centers when <see cref="LineKind"/> is <see cref="OfficeBorderLineKind.Double"/>.</summary>
    public double DoubleLineSeparation { get; }

    /// <summary>Whether this side should be drawn.</summary>
    public bool IsVisible => Width > 0D && Color.A > 0;

    /// <inheritdoc />
    public bool Equals(OfficeBorderSide other) =>
        Color == other.Color &&
        Width.Equals(other.Width) &&
        DashStyle == other.DashStyle &&
        LineKind == other.LineKind &&
        DoubleLineSeparation.Equals(other.DoubleLineSeparation);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeBorderSide other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            hash = (hash * 31) + Color.GetHashCode();
            hash = (hash * 31) + Width.GetHashCode();
            hash = (hash * 31) + DashStyle.GetHashCode();
            hash = (hash * 31) + LineKind.GetHashCode();
            hash = (hash * 31) + DoubleLineSeparation.GetHashCode();
            return hash;
        }
    }

    /// <summary>Compares two border sides for equality.</summary>
    public static bool operator ==(OfficeBorderSide left, OfficeBorderSide right) => left.Equals(right);

    /// <summary>Compares two border sides for inequality.</summary>
    public static bool operator !=(OfficeBorderSide left, OfficeBorderSide right) => !left.Equals(right);
}
