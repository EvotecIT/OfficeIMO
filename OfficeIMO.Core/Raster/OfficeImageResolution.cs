using System;

namespace OfficeIMO.Drawing;

/// <summary>Immutable native image resolution values and their physical unit or aspect-ratio unit.</summary>
public sealed class OfficeImageResolution {
    /// <summary>Creates a resolution using finite, positive values and a defined unit.</summary>
    public OfficeImageResolution(double horizontal, double vertical, OfficeImageResolutionUnit unit = OfficeImageResolutionUnit.PixelsPerInch) {
        if (horizontal <= 0 || double.IsNaN(horizontal) || double.IsInfinity(horizontal)) throw new ArgumentOutOfRangeException(nameof(horizontal));
        if (vertical <= 0 || double.IsNaN(vertical) || double.IsInfinity(vertical)) throw new ArgumentOutOfRangeException(nameof(vertical));
        if (unit < OfficeImageResolutionUnit.AspectRatio || unit > OfficeImageResolutionUnit.PixelsPerMeter) throw new ArgumentOutOfRangeException(nameof(unit));
        Horizontal = horizontal;
        Vertical = vertical;
        Unit = unit;
    }
    /// <summary>Horizontal resolution in <see cref="Unit"/>.</summary>
    public double Horizontal { get; }
    /// <summary>Vertical resolution in <see cref="Unit"/>.</summary>
    public double Vertical { get; }
    /// <summary>The native physical unit, or unitless aspect ratio.</summary>
    public OfficeImageResolutionUnit Unit { get; }
    /// <summary>Horizontal physical DPI, or null for a unitless aspect ratio.</summary>
    public double? PhysicalDpiX => ToDpi(Horizontal);
    /// <summary>Vertical physical DPI, or null for a unitless aspect ratio.</summary>
    public double? PhysicalDpiY => ToDpi(Vertical);
    private double? ToDpi(double value) => Unit switch {
        OfficeImageResolutionUnit.AspectRatio => null,
        OfficeImageResolutionUnit.PixelsPerCentimeter => value * 2.54D,
        OfficeImageResolutionUnit.PixelsPerMeter => value * 0.0254D,
        _ => value
    };
}
