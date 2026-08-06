using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free shadow intent for reusable drawing shapes.
/// Positive Y offsets move the shadow down in the shape's local top-left coordinate space.
/// </summary>
public sealed class OfficeShadow {
    /// <summary>Shadow color.</summary>
    public OfficeColor Color { get; }

    /// <summary>Shadow opacity from 0.0 to 1.0.</summary>
    public double Opacity { get; }

    /// <summary>Horizontal shadow offset in the caller's layout unit.</summary>
    public double OffsetX { get; }

    /// <summary>Vertical shadow offset in the caller's layout unit.</summary>
    public double OffsetY { get; }

    /// <summary>Approximate blur radius in the caller's layout unit.</summary>
    public double BlurRadius { get; }

    /// <summary>Creates a shadow descriptor.</summary>
    public OfficeShadow(OfficeColor color, double opacity, double offsetX, double offsetY, double blurRadius = 0D) {
        ValidateOpacity(opacity, nameof(opacity));
        ValidateFinite(offsetX, nameof(offsetX));
        ValidateFinite(offsetY, nameof(offsetY));
        ValidateNonNegativeFinite(blurRadius, nameof(blurRadius));

        Color = color;
        Opacity = opacity;
        OffsetX = offsetX;
        OffsetY = offsetY;
        BlurRadius = blurRadius;
    }

    /// <summary>Creates a detached copy.</summary>
    public OfficeShadow Clone() => new OfficeShadow(Color, Opacity, OffsetX, OffsetY, BlurRadius);

    private static void ValidateOpacity(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > 1D) {
            throw new ArgumentOutOfRangeException(paramName, "Shadow opacity must be a finite number between 0 and 1.");
        }
    }

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Shadow offsets must be finite numbers.");
        }
    }

    private static void ValidateNonNegativeFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Shadow blur radius must be a finite non-negative number.");
        }
    }
}
