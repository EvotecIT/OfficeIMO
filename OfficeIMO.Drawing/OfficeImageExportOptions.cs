using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared options for dependency-free image export pipelines.
/// </summary>
/// <remarks>
/// Document packages should inherit from this type for format-neutral settings and keep
/// document-specific layout policy in their own option types.
/// </remarks>
public class OfficeImageExportOptions {
    /// <summary>
    /// Output scale multiplier. A value of 2 creates a 2x raster or SVG surface.
    /// </summary>
    public double Scale { get; set; } = 1D;

    /// <summary>
    /// Background color used behind rendered document content.
    /// </summary>
    public OfficeColor BackgroundColor { get; set; } = OfficeColor.White;

    /// <summary>
    /// Format-specific settings used when the selected output is raster-based.
    /// </summary>
    public OfficeRasterEncodingOptions RasterEncoding { get; set; } = new OfficeRasterEncodingOptions();

    /// <summary>
    /// Validates that an export scale is finite and positive.
    /// </summary>
    /// <param name="scale">Scale value to validate.</param>
    /// <param name="paramName">Parameter name used for the thrown exception.</param>
    public static void ValidateScale(double scale, string paramName = "scale") {
        if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) {
            throw new ArgumentOutOfRangeException(paramName, "Scale must be a finite positive number.");
        }
    }
}
