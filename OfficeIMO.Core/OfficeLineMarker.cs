using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes a reusable marker drawn at the start or end of an open line.
/// </summary>
public sealed class OfficeLineMarker {
    /// <summary>Creates a line marker.</summary>
    public OfficeLineMarker(OfficeLineMarkerKind kind, double width, double length) {
        if (kind == OfficeLineMarkerKind.None) {
            throw new ArgumentOutOfRangeException(nameof(kind), "Use null instead of a None line marker.");
        }

        ValidatePositiveFinite(width, nameof(width));
        ValidatePositiveFinite(length, nameof(length));

        Kind = kind;
        Width = width;
        Length = length;
    }

    /// <summary>Marker shape.</summary>
    public OfficeLineMarkerKind Kind { get; }

    /// <summary>Marker width in the caller's layout unit.</summary>
    public double Width { get; }

    /// <summary>Marker length in the caller's layout unit.</summary>
    public double Length { get; }

    /// <summary>Creates a detached copy of this marker.</summary>
    public OfficeLineMarker Clone() => new OfficeLineMarker(Kind, Width, Length);

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Line marker dimensions must be finite positive numbers.");
        }
    }
}
