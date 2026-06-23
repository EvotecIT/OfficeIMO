using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Represents a physical page size used by image/page composition.
/// </summary>
public readonly struct OfficePageSize {
    /// <summary>
    /// Creates a physical page size measured in inches.
    /// </summary>
    public OfficePageSize(double widthInches, double heightInches) {
        ValidatePositive(widthInches, nameof(widthInches));
        ValidatePositive(heightInches, nameof(heightInches));
        WidthInches = widthInches;
        HeightInches = heightInches;
    }

    /// <summary>Page width in inches.</summary>
    public double WidthInches { get; }

    /// <summary>Page height in inches.</summary>
    public double HeightInches { get; }

    /// <summary>
    /// Creates a page size from millimeters.
    /// </summary>
    public static OfficePageSize FromMillimeters(double widthMillimeters, double heightMillimeters) {
        ValidatePositive(widthMillimeters, nameof(widthMillimeters));
        ValidatePositive(heightMillimeters, nameof(heightMillimeters));
        return new OfficePageSize(widthMillimeters / 25.4D, heightMillimeters / 25.4D);
    }

    /// <summary>
    /// Returns this page size in portrait orientation.
    /// </summary>
    public OfficePageSize Portrait() =>
        WidthInches <= HeightInches ? this : new OfficePageSize(HeightInches, WidthInches);

    /// <summary>
    /// Returns this page size in landscape orientation.
    /// </summary>
    public OfficePageSize Landscape() =>
        WidthInches >= HeightInches ? this : new OfficePageSize(HeightInches, WidthInches);

    /// <summary>
    /// Converts the page width to pixels for the requested DPI and scale.
    /// </summary>
    public int ToPixelWidth(double dpi, double scale = 1D) =>
        ToPixels(WidthInches, dpi, scale, nameof(dpi));

    /// <summary>
    /// Converts the page height to pixels for the requested DPI and scale.
    /// </summary>
    public int ToPixelHeight(double dpi, double scale = 1D) =>
        ToPixels(HeightInches, dpi, scale, nameof(dpi));

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
