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
    /// <summary>Default maximum number of pixels allocated for one raster export.</summary>
    public const long DefaultMaximumRasterPixels = 50_000_000L;

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
    /// Maximum number of pixels allocated for one raster export. Encoder-specific limits may reduce
    /// the effective ceiling further.
    /// </summary>
    public long MaximumRasterPixels { get; set; } = DefaultMaximumRasterPixels;

    /// <summary>
    /// Controls whether an oversized raster request is reduced to a safe scale or rejected.
    /// </summary>
    public OfficeRasterOverflowBehavior RasterOverflowBehavior { get; set; } = OfficeRasterOverflowBehavior.ReduceScale;

    /// <summary>
    /// Optional decoder for embedded source-image formats not handled by the dependency-free Drawing core.
    /// </summary>
    public IOfficeRasterImageCodec? ImageCodec { get; set; }

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

    /// <summary>Copies the shared image-export settings to another options instance.</summary>
    protected internal T CopyImageExportOptionsTo<T>(T target) where T : OfficeImageExportOptions {
        if (target == null) throw new ArgumentNullException(nameof(target));
        target.Scale = Scale;
        target.BackgroundColor = BackgroundColor;
        target.RasterEncoding = RasterEncoding?.Clone() ?? new OfficeRasterEncodingOptions();
        target.MaximumRasterPixels = MaximumRasterPixels;
        target.RasterOverflowBehavior = RasterOverflowBehavior;
        target.ImageCodec = ImageCodec;
        return target;
    }

    /// <summary>Validates the shared image-export settings.</summary>
    protected internal void ValidateImageExportOptions() {
        ValidateScale(Scale, nameof(Scale));
        if (MaximumRasterPixels < 1L) {
            throw new ArgumentOutOfRangeException(nameof(MaximumRasterPixels), "Maximum raster pixels must be positive.");
        }
        if (!Enum.IsDefined(typeof(OfficeRasterOverflowBehavior), RasterOverflowBehavior)) {
            throw new ArgumentOutOfRangeException(nameof(RasterOverflowBehavior));
        }
        if (RasterEncoding == null) {
            throw new InvalidOperationException("Raster encoding options cannot be null.");
        }
    }
}
