using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Represents a physical page size used by image/page composition.
/// </summary>
public readonly struct OfficePageSize {
    /// <summary>
    /// Creates a physical page size from positive document lengths.
    /// </summary>
    public OfficePageSize(OfficeLength width, OfficeLength height) {
        if (width <= OfficeLength.Zero) {
            throw new ArgumentOutOfRangeException(nameof(width), "Value must be positive.");
        }
        if (height <= OfficeLength.Zero) {
            throw new ArgumentOutOfRangeException(nameof(height), "Value must be positive.");
        }
        Width = width;
        Height = height;
    }

    /// <summary>Physical page width.</summary>
    public OfficeLength Width { get; }

    /// <summary>Physical page height.</summary>
    public OfficeLength Height { get; }

    /// <summary>
    /// Returns this page size in portrait orientation.
    /// </summary>
    public OfficePageSize Portrait() =>
        Width <= Height ? this : new OfficePageSize(Height, Width);

    /// <summary>
    /// Returns this page size in landscape orientation.
    /// </summary>
    public OfficePageSize Landscape() =>
        Width >= Height ? this : new OfficePageSize(Height, Width);

    /// <summary>
    /// Converts the page width to pixels for the requested DPI and scale.
    /// </summary>
    public int ToPixelWidth(double dpi, double scale = 1D) =>
        ToPixels(Width.ToInches(), dpi, scale, nameof(dpi));

    /// <summary>
    /// Converts the page height to pixels for the requested DPI and scale.
    /// </summary>
    public int ToPixelHeight(double dpi, double scale = 1D) =>
        ToPixels(Height.ToInches(), dpi, scale, nameof(dpi));

    private static int ToPixels(double inches, double dpi, double scale, string dpiParameterName) {
        ValidatePositive(dpi, dpiParameterName);
        ValidatePositive(scale, nameof(scale));
        return Math.Max(1, (int)Math.Ceiling(inches * dpi * scale));
    }

    private static void ValidatePositive(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new ArgumentOutOfRangeException(parameterName, "Value must be positive.");
        }
    }
}
