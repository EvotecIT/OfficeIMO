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
    /// Optional maximum output width in pixels (or CSS pixels for SVG). The requested scale is
    /// reduced as needed while preserving aspect ratio; smaller outputs are never enlarged.
    /// </summary>
    public int? MaximumOutputWidth { get; set; }

    /// <summary>
    /// Optional maximum output height in pixels (or CSS pixels for SVG). The requested scale is
    /// reduced as needed while preserving aspect ratio; smaller outputs are never enlarged.
    /// </summary>
    public int? MaximumOutputHeight { get; set; }

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

    /// <summary>
    /// Optional host text shaper used for deterministic complex-script raster output.
    /// </summary>
    /// <remarks>
    /// Hosts can adapt HarfBuzz, DirectWrite, Core Text, or another shaping engine without adding
    /// that dependency to OfficeIMO. When batch concurrency is enabled, the provider must be safe
    /// for concurrent calls. Returning <see langword="null"/> keeps the managed fallback path.
    /// </remarks>
    public IOfficeTextShapingProvider? TextShapingProvider { get; set; }

    /// <summary>Optional BCP 47 language hint passed to <see cref="TextShapingProvider"/>.</summary>
    public string? TextShapingLanguage { get; set; }

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
    /// Maximum duration allowed for one render operation. The default is unlimited.
    /// Cancellation-aware renderers are interrupted at the deadline; legacy renderers that do not
    /// observe cancellation are rejected when they return after the deadline.
    /// </summary>
    public TimeSpan RenderTimeout { get; set; } = System.Threading.Timeout.InfiniteTimeSpan;

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
        target.MaximumOutputWidth = MaximumOutputWidth;
        target.MaximumOutputHeight = MaximumOutputHeight;
        target.BackgroundColor = BackgroundColor;
        target.RasterEncoding = RasterEncoding?.Clone() ?? new OfficeRasterEncodingOptions();
        target.MaximumRasterPixels = MaximumRasterPixels;
        target.RasterOverflowBehavior = RasterOverflowBehavior;
        target.ImageCodec = ImageCodec;
        target.TargetDpi = TargetDpi;
        target.Fonts = Fonts?.Clone() ?? new OfficeFontFaceCollection();
        target.TextShapingProvider = TextShapingProvider;
        target.TextShapingLanguage = TextShapingLanguage;
        target.Policy = Policy?.Clone() ?? new OfficeImageExportPolicy();
        target.Progress = Progress;
        target.MaximumOutputCount = MaximumOutputCount;
        target.MaximumTotalRasterPixels = MaximumTotalRasterPixels;
        target.MaximumTotalEncodedBytes = MaximumTotalEncodedBytes;
        target.RenderTimeout = RenderTimeout;
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
        if (MaximumOutputWidth.HasValue && MaximumOutputWidth.Value < 1) {
            throw new ArgumentOutOfRangeException(nameof(MaximumOutputWidth));
        }
        if (MaximumOutputHeight.HasValue && MaximumOutputHeight.Value < 1) {
            throw new ArgumentOutOfRangeException(nameof(MaximumOutputHeight));
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
        TextShapingLanguage = string.IsNullOrWhiteSpace(TextShapingLanguage)
            ? null
            : TextShapingLanguage!.Trim();
        if (Policy == null) throw new InvalidOperationException("Image export policy cannot be null.");
        if (MaximumOutputCount < 1) throw new ArgumentOutOfRangeException(nameof(MaximumOutputCount));
        if (MaximumTotalRasterPixels < 1L) throw new ArgumentOutOfRangeException(nameof(MaximumTotalRasterPixels));
        if (MaximumTotalEncodedBytes < 1L) throw new ArgumentOutOfRangeException(nameof(MaximumTotalEncodedBytes));
        if (RenderTimeout != System.Threading.Timeout.InfiniteTimeSpan &&
            (RenderTimeout <= TimeSpan.Zero || RenderTimeout.TotalMilliseconds > int.MaxValue)) {
            throw new ArgumentOutOfRangeException(
                nameof(RenderTimeout),
                "Render timeout must be positive, no greater than Int32.MaxValue milliseconds, or infinite.");
        }
        if (MaximumDegreeOfParallelism < 1) throw new ArgumentOutOfRangeException(nameof(MaximumDegreeOfParallelism));
    }

    private static void ValidateDpi(double value, string name) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value) || value > ushort.MaxValue) {
            throw new ArgumentOutOfRangeException(name, "Raster DPI must be finite, positive, and encodable by every shared raster format.");
        }
    }

    /// <summary>
    /// Resolves the requested scale after applying optional output dimension caps while preserving aspect ratio.
    /// </summary>
    public double GetEffectiveScale(double logicalWidth, double logicalHeight) {
        if (logicalWidth <= 0D || double.IsNaN(logicalWidth) || double.IsInfinity(logicalWidth)) {
            throw new ArgumentOutOfRangeException(nameof(logicalWidth));
        }
        if (logicalHeight <= 0D || double.IsNaN(logicalHeight) || double.IsInfinity(logicalHeight)) {
            throw new ArgumentOutOfRangeException(nameof(logicalHeight));
        }

        double scale = TargetDpi.HasValue ? TargetDpi.Value / LogicalUnitsPerInch : Scale;
        if (MaximumOutputWidth.HasValue) {
            scale = Math.Min(scale, GetCeilingSafeMaximumScale(logicalWidth, MaximumOutputWidth.Value));
        }
        if (MaximumOutputHeight.HasValue) {
            scale = Math.Min(scale, GetCeilingSafeMaximumScale(logicalHeight, MaximumOutputHeight.Value));
        }
        if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) {
            throw new ArgumentOutOfRangeException(nameof(logicalWidth), "The requested output bounds cannot be represented by a finite positive scale.");
        }
        return scale;
    }

    private static double GetCeilingSafeMaximumScale(double logicalDimension, int maximumDimension) {
        double scale = maximumDimension / logicalDimension;
        while (scale > 0D && Math.Ceiling(logicalDimension * scale) > maximumDimension) {
            long bits = BitConverter.DoubleToInt64Bits(scale);
            if (bits <= 0L) return 0D;
            scale = BitConverter.Int64BitsToDouble(bits - 1L);
        }
        return scale;
    }
}
