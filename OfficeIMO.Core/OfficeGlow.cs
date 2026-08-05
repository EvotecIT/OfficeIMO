using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free glow intent for reusable drawing shapes.
/// Renderers may approximate the glow with layered translucent strokes.
/// </summary>
public sealed class OfficeGlow {
    /// <summary>Glow color.</summary>
    public OfficeColor Color { get; }

    /// <summary>Glow opacity from 0.0 to 1.0.</summary>
    public double Opacity { get; }

    /// <summary>Glow radius in the caller's layout unit.</summary>
    public double Radius { get; }

    /// <summary>Creates a glow descriptor.</summary>
    public OfficeGlow(OfficeColor color, double opacity, double radius) {
        ValidateOpacity(opacity, nameof(opacity));
        ValidateRadius(radius, nameof(radius));

        Color = color;
        Opacity = opacity;
        Radius = radius;
    }

    /// <summary>Creates a detached copy.</summary>
    public OfficeGlow Clone() => new OfficeGlow(Color, Opacity, Radius);

    private static void ValidateOpacity(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > 1D) {
            throw new ArgumentOutOfRangeException(paramName, "Glow opacity must be a finite number between 0 and 1.");
        }
    }

    private static void ValidateRadius(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Glow radius must be a finite non-negative number.");
        }
    }
}
