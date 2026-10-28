using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared fluent image-export surface for document-level batch exports.
/// </summary>
/// <typeparam name="TBuilder">Concrete builder type returned from fluent methods.</typeparam>
/// <typeparam name="TOptions">Document-specific image export options.</typeparam>
public abstract class OfficeImageExportBatchBuilder<TBuilder, TOptions>
    where TBuilder : OfficeImageExportBatchBuilder<TBuilder, TOptions>
    where TOptions : OfficeImageExportOptions {
    private readonly Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> _export;
    private readonly Func<OfficeImageExportFormat, TOptions, CancellationToken, Task<IReadOnlyList<OfficeImageExportResult>>>? _exportAsync;
    private readonly Action<OfficeImageExportFormat, TOptions, OfficeImageExportConsumer, CancellationToken>? _exportEach;
    private readonly Func<OfficeImageExportFormat, TOptions, OfficeImageExportAsyncConsumer, CancellationToken, Task>? _exportEachAsync;
    private OfficeImageExportFormat _format = OfficeImageExportFormat.Png;
    private OfficeImageExportFileConflictPolicy _conflictPolicy = OfficeImageExportFileConflictPolicy.FailIfExists;

    /// <summary>Creates a batch builder over an existing materializing export function.</summary>
    protected OfficeImageExportBatchBuilder(
        TOptions options,
        Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> export) {
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _export = export ?? throw new ArgumentNullException(nameof(export));
    }

    /// <summary>Creates a batch builder with a genuine asynchronous materializing renderer.</summary>
    protected OfficeImageExportBatchBuilder(
        TOptions options,
        Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> export,
        Func<OfficeImageExportFormat, TOptions, CancellationToken, Task<IReadOnlyList<OfficeImageExportResult>>> exportAsync)
        : this(options, export) {
        _exportAsync = exportAsync ?? throw new ArgumentNullException(nameof(exportAsync));
    }

    /// <summary>Creates a batch builder with synchronous streaming support.</summary>
    protected OfficeImageExportBatchBuilder(
        TOptions options,
        Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> export,
        Action<OfficeImageExportFormat, TOptions, OfficeImageExportConsumer, CancellationToken> exportEach)
        : this(options, export) {
        _exportEach = exportEach ?? throw new ArgumentNullException(nameof(exportEach));
    }

    /// <summary>Creates a batch builder with synchronous and asynchronous streaming support.</summary>
    protected OfficeImageExportBatchBuilder(
        TOptions options,
        Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> export,
        Action<OfficeImageExportFormat, TOptions, OfficeImageExportConsumer, CancellationToken> exportEach,
        Func<OfficeImageExportFormat, TOptions, OfficeImageExportAsyncConsumer, CancellationToken, Task> exportEachAsync)
        : this(options, export, exportEach) {
        _exportEachAsync = exportEachAsync ?? throw new ArgumentNullException(nameof(exportEachAsync));
    }

    /// <summary>Document-specific options being configured by this builder.</summary>
    protected TOptions Options { get; }

    /// <summary>
    /// Configures the complete document-specific option set for settings that do not have a dedicated fluent shortcut.
    /// </summary>
    public TBuilder ConfigureOptions(Action<TOptions> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        configure(Options);
        return This;
    }

    /// <summary>Configures PNG output.</summary>
    public TBuilder AsPng() => As(OfficeImageExportFormat.Png);

    /// <summary>Configures SVG output.</summary>
    public TBuilder AsSvg() => As(OfficeImageExportFormat.Svg);

    /// <summary>Configures JPEG output.</summary>
    public TBuilder AsJpeg() => As(OfficeImageExportFormat.Jpeg);

    /// <summary>Configures TIFF output.</summary>
    public TBuilder AsTiff() => As(OfficeImageExportFormat.Tiff);

    /// <summary>Configures lossless WebP output.</summary>
    public TBuilder AsWebp() => As(OfficeImageExportFormat.Webp);

    /// <summary>Configures the output image format.</summary>
    public TBuilder As(OfficeImageExportFormat format) {
        if (!Enum.IsDefined(typeof(OfficeImageExportFormat), format)) throw new ArgumentOutOfRangeException(nameof(format));
        _format = format;
        return This;
    }

    /// <summary>Sets the output scale.</summary>
    public TBuilder WithScale(double scale) {
        OfficeImageExportOptions.ValidateScale(scale);
        Options.Scale = scale;
        Options.TargetDpi = null;
        return This;
    }

    /// <summary>Fits every output within the specified pixel dimensions without changing aspect ratio or enlarging it.</summary>
    public TBuilder FitWithin(int maximumWidth, int maximumHeight) {
        if (maximumWidth < 1) throw new ArgumentOutOfRangeException(nameof(maximumWidth));
        if (maximumHeight < 1) throw new ArgumentOutOfRangeException(nameof(maximumHeight));
        Options.MaximumOutputWidth = maximumWidth;
        Options.MaximumOutputHeight = maximumHeight;
        return This;
    }

    /// <summary>Fits every output within the specified pixel width without enlarging it.</summary>
    public TBuilder FitWithinWidth(int maximumWidth) {
        if (maximumWidth < 1) throw new ArgumentOutOfRangeException(nameof(maximumWidth));
        Options.MaximumOutputWidth = maximumWidth;
        return This;
    }

    /// <summary>Fits every output within the specified pixel height without enlarging it.</summary>
    public TBuilder FitWithinHeight(int maximumHeight) {
        if (maximumHeight < 1) throw new ArgumentOutOfRangeException(nameof(maximumHeight));
        Options.MaximumOutputHeight = maximumHeight;
        return This;
    }

    /// <summary>Removes output dimension caps.</summary>
    public TBuilder WithoutSizeLimit() {
        Options.MaximumOutputWidth = null;
        Options.MaximumOutputHeight = null;
        return This;
    }

    /// <summary>Sets a target output density and lets the document adapter resolve its logical-unit scale.</summary>
    public TBuilder AtDpi(double dpi) {
        if (dpi <= 0D || double.IsNaN(dpi) || double.IsInfinity(dpi)) throw new ArgumentOutOfRangeException(nameof(dpi));
        Options.TargetDpi = dpi;
        return This;
    }

    /// <summary>Sets the export background color.</summary>
    public TBuilder WithBackground(OfficeColor color) {
        Options.BackgroundColor = color;
        return This;
    }

    /// <summary>Sets the export background from a named color or hexadecimal color value.</summary>
    public TBuilder WithBackground(string color) => WithBackground(OfficeColor.Parse(color));

    /// <summary>Configures format-specific raster encoding settings.</summary>
    public TBuilder WithRasterEncoding(Action<OfficeRasterEncodingOptions> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        OfficeRasterEncodingOptions settings = Options.RasterEncoding ?? new OfficeRasterEncodingOptions();
        configure(settings);
        Options.RasterEncoding = settings;
        return This;
    }

    /// <summary>Sets the maximum pixel allocation for each raster result.</summary>
    public TBuilder WithMaximumRasterPixels(long maximumPixels) {
        if (maximumPixels < 1L) throw new ArgumentOutOfRangeException(nameof(maximumPixels));
        Options.MaximumRasterPixels = maximumPixels;
        return This;
    }

    /// <summary>Sets aggregate count, raster-pixel, and encoded-byte budgets for one batch.</summary>
    public TBuilder WithBatchLimits(int maximumOutputCount, long maximumTotalRasterPixels, long maximumTotalEncodedBytes) {
        if (maximumOutputCount < 1) throw new ArgumentOutOfRangeException(nameof(maximumOutputCount));
        if (maximumTotalRasterPixels < 1L) throw new ArgumentOutOfRangeException(nameof(maximumTotalRasterPixels));
        if (maximumTotalEncodedBytes < 1L) throw new ArgumentOutOfRangeException(nameof(maximumTotalEncodedBytes));
        Options.MaximumOutputCount = maximumOutputCount;
        Options.MaximumTotalRasterPixels = maximumTotalRasterPixels;
        Options.MaximumTotalEncodedBytes = maximumTotalEncodedBytes;
        return This;
    }

    /// <summary>Sets the maximum duration allowed for one batch render operation.</summary>
    public TBuilder WithRenderTimeout(TimeSpan timeout) {
        if (timeout != System.Threading.Timeout.InfiniteTimeSpan &&
            (timeout <= TimeSpan.Zero || timeout.TotalMilliseconds > int.MaxValue)) {
            throw new ArgumentOutOfRangeException(nameof(timeout));
        }
        Options.RenderTimeout = timeout;
        return This;
    }

    /// <summary>
    /// Enables bounded parallel rendering for adapters whose selected items are independent.
    /// Output order remains deterministic.
    /// </summary>
    public TBuilder WithMaximumConcurrency(int maximumDegreeOfParallelism) {
        if (maximumDegreeOfParallelism < 1) throw new ArgumentOutOfRangeException(nameof(maximumDegreeOfParallelism));
        Options.MaximumDegreeOfParallelism = maximumDegreeOfParallelism;
        return This;
    }

    /// <summary>Sets the policy applied when requested raster dimensions exceed a safety limit.</summary>
    public TBuilder OnRasterOverflow(OfficeRasterOverflowBehavior behavior) {
        if (!Enum.IsDefined(typeof(OfficeRasterOverflowBehavior), behavior)) throw new ArgumentOutOfRangeException(nameof(behavior));
        Options.RasterOverflowBehavior = behavior;
        return This;
    }

    /// <summary>Sets the optional decoder used for embedded source-image formats outside Drawing's built-in set.</summary>
    public TBuilder WithImageCodec(IOfficeRasterImageCodec imageCodec) {
        Options.ImageCodec = imageCodec ?? throw new ArgumentNullException(nameof(imageCodec));
        return This;
    }

    /// <summary>Adds a deterministic caller-supplied TrueType face.</summary>
    public TBuilder WithFont(string familyName, byte[] data, OfficeFontStyle style = OfficeFontStyle.Regular) {
        Options.Fonts.Add(familyName, data, style);
        return This;
    }

    /// <summary>Adds deterministic caller-supplied TrueType faces.</summary>
    public TBuilder WithFonts(OfficeFontFaceCollection fonts) {
        Options.Fonts.AddRange(fonts ?? throw new ArgumentNullException(nameof(fonts)));
        return This;
    }

    /// <summary>Uses a host shaping engine for complex-script raster text.</summary>
    public TBuilder WithTextShaping(IOfficeTextShapingProvider provider, string? language = null) {
        Options.TextShapingProvider = provider ?? throw new ArgumentNullException(nameof(provider));
        Options.TextShapingLanguage = string.IsNullOrWhiteSpace(language) ? null : language!.Trim();
        return This;
    }

    /// <summary>Configures diagnostic acceptance before results are returned or committed.</summary>
    public TBuilder WithPolicy(Action<OfficeImageExportPolicy> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        configure(Options.Policy);
        return This;
    }

    /// <summary>Sets a progress observer for render and save stages.</summary>
    public TBuilder WithProgress(IProgress<OfficeImageExportProgress> progress) {
        Options.Progress = progress ?? throw new ArgumentNullException(nameof(progress));
        return This;
    }

    /// <summary>Sets how file saves handle existing destinations.</summary>
    public TBuilder OnFileConflict(OfficeImageExportFileConflictPolicy policy) {
        if (!Enum.IsDefined(typeof(OfficeImageExportFileConflictPolicy), policy)) throw new ArgumentOutOfRangeException(nameof(policy));
        _conflictPolicy = policy;
        return This;
    }

    /// <summary>Configures a standard preview profile: PNG, 1x scale, white background.</summary>
    public TBuilder ForPreview() {
        _format = OfficeImageExportFormat.Png;
        Options.Scale = 1D;
        Options.TargetDpi = null;
        Options.BackgroundColor = OfficeColor.White;
        return This;
    }

    /// <summary>Selects a shared output-density preset while retaining format, fonts and safety limits.</summary>
    public TBuilder WithQuality(OfficeImageExportQuality quality) {
        Options.UseQuality(quality);
        return This;
    }

    /// <summary>Configures a print profile with an explicit target DPI.</summary>
    public TBuilder ForPrint(double dpi = 300D) {
        if (dpi <= 0D || double.IsNaN(dpi) || double.IsInfinity(dpi)) throw new ArgumentOutOfRangeException(nameof(dpi));
        _format = OfficeImageExportFormat.Png;
        Options.TargetDpi = dpi;
        Options.BackgroundColor = OfficeColor.White;
        return This;
    }

    private OfficeImageExportBatchOperation<TOptions> Operation => new OfficeImageExportBatchOperation<TOptions>(
        Options, _format, _conflictPolicy, _export, _exportAsync, _exportEach, _exportEachAsync);

    /// <summary>Exports all selected images and materializes their encoded payloads.</summary>
    public IReadOnlyList<OfficeImageExportResult> Export() => Operation.Export();

    /// <summary>Asynchronously exports all selected images and materializes their encoded payloads.</summary>
    public Task<IReadOnlyList<OfficeImageExportResult>> ExportAsync(CancellationToken cancellationToken = default) => Operation.ExportAsync(cancellationToken);

    /// <summary>Streams results without retaining their encoded payloads.</summary>
    public void ExportEach(OfficeImageExportConsumer consumer, CancellationToken cancellationToken = default) => Operation.ExportEach(consumer, cancellationToken);

    /// <summary>Streams results asynchronously without retaining their encoded payloads.</summary>
    public Task ExportEachAsync(OfficeImageExportAsyncConsumer consumer, CancellationToken cancellationToken = default) => Operation.ExportEachAsync(consumer, cancellationToken);

    /// <summary>Saves images and returns payload-bearing results with normalized paths.</summary>
    public IReadOnlyList<OfficeImageExportResult> Save(string folderPath) => Operation.Save(folderPath);

    /// <summary>Saves images without retaining encoded payloads in the returned result.</summary>
    public OfficeImageExportBatchSaveResult SaveFiles(string folderPath) => Operation.SaveFiles(folderPath);

    /// <summary>Asynchronously saves images and returns payload-bearing results with normalized paths.</summary>
    public Task<IReadOnlyList<OfficeImageExportResult>> SaveAsync(string folderPath, CancellationToken cancellationToken = default) => Operation.SaveAsync(folderPath, cancellationToken);

    /// <summary>Asynchronously saves images without retaining encoded payloads in the returned result.</summary>
    public Task<OfficeImageExportBatchSaveResult> SaveFilesAsync(string folderPath, CancellationToken cancellationToken = default) => Operation.SaveFilesAsync(folderPath, cancellationToken);

    private TBuilder This => (TBuilder)this;
}
