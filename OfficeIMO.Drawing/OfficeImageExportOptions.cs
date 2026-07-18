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

    /// <summary>Default maximum number of images produced by one batch operation.</summary>
    public const int DefaultMaximumOutputCount = 10_000;

    /// <summary>Default maximum aggregate raster pixels produced by one batch operation.</summary>
    public const long DefaultMaximumTotalRasterPixels = 500_000_000L;

    /// <summary>Default maximum aggregate encoded bytes produced by one batch operation.</summary>
    public const long DefaultMaximumTotalEncodedBytes = 1024L * 1024L * 1024L;

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

    /// <summary>Optional target output density. Document adapters map logical units to this DPI.</summary>
    public double? TargetDpi { get; set; }

    /// <summary>Caller-supplied deterministic TrueType faces used before platform fallback.</summary>
    public OfficeFontFaceCollection Fonts { get; set; } = new OfficeFontFaceCollection();

    /// <summary>Diagnostic acceptance policy applied before an export is returned or committed.</summary>
    public OfficeImageExportPolicy Policy { get; set; } = new OfficeImageExportPolicy();

    /// <summary>Optional progress observer for rendering and saving.</summary>
    public IProgress<OfficeImageExportProgress>? Progress { get; set; }

    /// <summary>Maximum number of results accepted from one batch export.</summary>
    public int MaximumOutputCount { get; set; } = DefaultMaximumOutputCount;

    /// <summary>Maximum aggregate raster pixels accepted from one batch export.</summary>
    public long MaximumTotalRasterPixels { get; set; } = DefaultMaximumTotalRasterPixels;

    /// <summary>Maximum aggregate encoded bytes accepted from one batch export.</summary>
    public long MaximumTotalEncodedBytes { get; set; } = DefaultMaximumTotalEncodedBytes;

    /// <summary>
    /// Maximum concurrent independent renders. Defaults to one because callers must opt in when their
    /// document model can be read concurrently.
    /// </summary>
    public int MaximumDegreeOfParallelism { get; set; } = 1;

    /// <summary>Logical document units represented by one inch. Point-based adapters override with 72.</summary>
    public virtual double LogicalUnitsPerInch => 96D;

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

    /// <summary>
    /// Validates this options snapshot and throws when an exported result violates its acceptance policy.
    /// Format packages call this at their direct export boundary so fluent and non-fluent APIs behave identically.
    /// </summary>
    public OfficeImageExportResult EnsureAccepted(OfficeImageExportResult result) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        ValidateImageExportOptions();
        Policy.EnsureAccepted(result.Diagnostics);
        return result;
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
        target.TargetDpi = TargetDpi;
        target.Fonts = Fonts?.Clone() ?? new OfficeFontFaceCollection();
        target.Policy = Policy?.Clone() ?? new OfficeImageExportPolicy();
        target.Progress = Progress;
        target.MaximumOutputCount = MaximumOutputCount;
        target.MaximumTotalRasterPixels = MaximumTotalRasterPixels;
        target.MaximumTotalEncodedBytes = MaximumTotalEncodedBytes;
        target.MaximumDegreeOfParallelism = MaximumDegreeOfParallelism;
        return target;
    }

    /// <summary>
    /// Creates the detached, validated snapshot consumed by one fluent export operation.
    /// Target-DPI-derived scale and density are intentionally resolved only on the snapshot.
    /// </summary>
    internal T CreateEffectiveImageExportOptions<T>() where T : OfficeImageExportOptions {
        var effective = (T)MemberwiseClone();
        effective.RasterEncoding = RasterEncoding?.Clone()!;
        effective.Fonts = Fonts?.Clone()!;
        effective.Policy = Policy?.Clone()!;
        effective.ValidateImageExportOptions();
        return effective;
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
        if (TargetDpi.HasValue &&
            (TargetDpi.Value <= 0D || double.IsNaN(TargetDpi.Value) || double.IsInfinity(TargetDpi.Value))) {
            throw new ArgumentOutOfRangeException(nameof(TargetDpi), "Target DPI must be finite and positive.");
        }
        if (TargetDpi.HasValue) {
            Scale = TargetDpi.Value / LogicalUnitsPerInch;
            RasterEncoding.DpiX = TargetDpi.Value;
            RasterEncoding.DpiY = TargetDpi.Value;
        }
        ValidateDpi(RasterEncoding.DpiX, nameof(RasterEncoding.DpiX));
        ValidateDpi(RasterEncoding.DpiY, nameof(RasterEncoding.DpiY));
        if (Fonts == null) throw new InvalidOperationException("Font collection cannot be null.");
        if (Policy == null) throw new InvalidOperationException("Image export policy cannot be null.");
        if (MaximumOutputCount < 1) throw new ArgumentOutOfRangeException(nameof(MaximumOutputCount));
        if (MaximumTotalRasterPixels < 1L) throw new ArgumentOutOfRangeException(nameof(MaximumTotalRasterPixels));
        if (MaximumTotalEncodedBytes < 1L) throw new ArgumentOutOfRangeException(nameof(MaximumTotalEncodedBytes));
        if (MaximumDegreeOfParallelism < 1) throw new ArgumentOutOfRangeException(nameof(MaximumDegreeOfParallelism));
    }

    private static void ValidateDpi(double value, string name) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value) || value > ushort.MaxValue) {
            throw new ArgumentOutOfRangeException(name, "Raster DPI must be finite, positive, and encodable by every shared raster format.");
        }
    }
}
