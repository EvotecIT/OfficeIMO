using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared fluent image-export surface used by OfficeIMO document packages.
/// </summary>
/// <typeparam name="TBuilder">Concrete builder type returned from fluent methods.</typeparam>
/// <typeparam name="TOptions">Document-specific image export options.</typeparam>
public abstract class OfficeImageExportBuilder<TBuilder, TOptions>
    where TBuilder : OfficeImageExportBuilder<TBuilder, TOptions>
    where TOptions : OfficeImageExportOptions {
    private readonly Func<OfficeImageExportFormat, TOptions, OfficeImageExportResult> _export;
    private readonly Func<OfficeImageExportFormat, TOptions, CancellationToken, OfficeImageExportResult>? _exportWithCancellation;
    private readonly Func<OfficeImageExportFormat, TOptions, CancellationToken, Task<OfficeImageExportResult>>? _exportAsync;
    private OfficeImageExportFormat _format = OfficeImageExportFormat.Png;
    private OfficeImageExportFileConflictPolicy _conflictPolicy = OfficeImageExportFileConflictPolicy.FailIfExists;

    /// <summary>
    /// Creates a fluent export builder over an existing document-specific export function.
    /// </summary>
    protected OfficeImageExportBuilder(TOptions options, Func<OfficeImageExportFormat, TOptions, OfficeImageExportResult> export) {
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _export = export ?? throw new ArgumentNullException(nameof(export));
    }

    /// <summary>Creates a builder over a cancellation-aware synchronous renderer.</summary>
    protected OfficeImageExportBuilder(
        TOptions options,
        Func<OfficeImageExportFormat, TOptions, CancellationToken, OfficeImageExportResult> export) {
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _exportWithCancellation = export ?? throw new ArgumentNullException(nameof(export));
        _export = (format, effective) => _exportWithCancellation(format, effective, CancellationToken.None);
    }

    /// <summary>
    /// Creates a fluent export builder with a genuine asynchronous renderer for resource-aware document models.
    /// </summary>
    protected OfficeImageExportBuilder(
        TOptions options,
        Func<OfficeImageExportFormat, TOptions, OfficeImageExportResult> export,
        Func<OfficeImageExportFormat, TOptions, CancellationToken, Task<OfficeImageExportResult>> exportAsync)
        : this(options, export) {
        _exportAsync = exportAsync ?? throw new ArgumentNullException(nameof(exportAsync));
    }

    /// <summary>Document-specific options being configured by this builder.</summary>
    protected TOptions Options { get; }

    /// <summary>Configures PNG output.</summary>
    public TBuilder AsPng() {
        _format = OfficeImageExportFormat.Png;
        return This;
    }

    /// <summary>Configures SVG output.</summary>
    public TBuilder AsSvg() {
        _format = OfficeImageExportFormat.Svg;
        return This;
    }

    /// <summary>Configures JPEG output.</summary>
    public TBuilder AsJpeg() {
        _format = OfficeImageExportFormat.Jpeg;
        return This;
    }

    /// <summary>Configures TIFF output.</summary>
    public TBuilder AsTiff() {
        _format = OfficeImageExportFormat.Tiff;
        return This;
    }

    /// <summary>Configures lossless WebP output.</summary>
    public TBuilder AsWebp() {
        _format = OfficeImageExportFormat.Webp;
        return This;
    }

    /// <summary>Configures the output image format.</summary>
    public TBuilder As(OfficeImageExportFormat format) {
        if (!Enum.IsDefined(typeof(OfficeImageExportFormat), format)) {
            throw new ArgumentOutOfRangeException(nameof(format));
        }

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

    /// <summary>Sets the maximum pixel allocation for one raster result.</summary>
    public TBuilder WithMaximumRasterPixels(long maximumPixels) {
        if (maximumPixels < 1L) throw new ArgumentOutOfRangeException(nameof(maximumPixels));
        Options.MaximumRasterPixels = maximumPixels;
        return This;
    }

    /// <summary>Sets the policy applied when requested raster dimensions exceed a safety limit.</summary>
    public TBuilder OnRasterOverflow(OfficeRasterOverflowBehavior behavior) {
        if (!Enum.IsDefined(typeof(OfficeRasterOverflowBehavior), behavior)) {
            throw new ArgumentOutOfRangeException(nameof(behavior));
        }
        Options.RasterOverflowBehavior = behavior;
        return This;
    }

    /// <summary>Sets the optional decoder used for embedded source-image formats outside Drawing's built-in set.</summary>
    public TBuilder WithImageCodec(IOfficeRasterImageCodec imageCodec) {
        Options.ImageCodec = imageCodec ?? throw new ArgumentNullException(nameof(imageCodec));
        return This;
    }

    /// <summary>Sets a target output density and lets the document adapter resolve its logical-unit scale.</summary>
    public TBuilder AtDpi(double dpi) {
        if (dpi <= 0D || double.IsNaN(dpi) || double.IsInfinity(dpi)) throw new ArgumentOutOfRangeException(nameof(dpi));
        Options.TargetDpi = dpi;
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

    /// <summary>Sets how file saves handle an existing destination.</summary>
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

    /// <summary>Configures a print profile with an explicit target DPI.</summary>
    public TBuilder ForPrint(double dpi = 300D) {
        if (dpi <= 0D || double.IsNaN(dpi) || double.IsInfinity(dpi)) throw new ArgumentOutOfRangeException(nameof(dpi));
        _format = OfficeImageExportFormat.Png;
        Options.TargetDpi = dpi;
        Options.BackgroundColor = OfficeColor.White;
        return This;
    }

    /// <summary>Exports using the currently configured format and options.</summary>
    public OfficeImageExportResult Export() => Export(CancellationToken.None);

    /// <summary>Exports using the currently configured format and observes cancellation during supported render stages.</summary>
    public OfficeImageExportResult Export(CancellationToken cancellationToken) {
        OfficeImageExportResult result = Render(cancellationToken);
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name));
        return result;
    }

    /// <summary>Exports asynchronously when the adapter owns a genuine asynchronous render path.</summary>
    public async Task<OfficeImageExportResult> ExportAsync(CancellationToken cancellationToken = default) {
        OfficeImageExportResult result = await RenderAsync(cancellationToken).ConfigureAwait(false);
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name));
        return result;
    }

    /// <summary>Exports and returns the encoded image bytes.</summary>
    public byte[] ToBytes() => Export().Bytes;

    /// <summary>Saves the exported image to a file path.</summary>
    public OfficeImageExportResult Save(string path) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));
        }

        string resolvedPath = OfficeImageExportPath.ResolveFile(path, _format, _conflictPolicy);
        OfficeImageExportResult result = Render(CancellationToken.None);
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, 0, 1, result.Name, resolvedPath));
        OfficeImageExportResult saved = result.Save(resolvedPath, _conflictPolicy);
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name, resolvedPath));
        return saved;
    }

    /// <summary>Writes the exported image to a stream.</summary>
    public OfficeImageExportResult Save(Stream stream) {
        OfficeImageExportResult result = Render(CancellationToken.None);
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, 0, 1, result.Name));
        OfficeImageExportResult saved = result.Save(stream);
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name));
        return saved;
    }

    /// <summary>Asynchronously saves the exported image to a file path.</summary>
    public async Task<OfficeImageExportResult> SaveAsync(
        string path,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));
        }

        cancellationToken.ThrowIfCancellationRequested();
        string resolvedPath = OfficeImageExportPath.ResolveFile(path, _format, _conflictPolicy);
        OfficeImageExportResult result = await RenderAsync(cancellationToken).ConfigureAwait(false);
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, 0, 1, result.Name, resolvedPath));
        OfficeImageExportResult saved = await result.SaveAsync(
            resolvedPath,
            _conflictPolicy,
            cancellationToken).ConfigureAwait(false);
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name, resolvedPath));
        return saved;
    }

    /// <summary>Asynchronously writes the exported image to a stream.</summary>
    public async Task<OfficeImageExportResult> SaveAsync(
        Stream stream,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeImageExportResult result = await RenderAsync(cancellationToken).ConfigureAwait(false);
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, 0, 1, result.Name));
        OfficeImageExportResult saved = await result.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Completed, 1, 1, result.Name));
        return saved;
    }

    private OfficeImageExportResult Render(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        Options.ValidateImageExportOptions();
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Rendering, 0, 1));
        OfficeImageExportResult result = _exportWithCancellation != null
            ? _exportWithCancellation(_format, Options, cancellationToken)
            : _export(_format, Options);
        cancellationToken.ThrowIfCancellationRequested();
        result.Require(Options.Policy);
        return result;
    }

    private async Task<OfficeImageExportResult> RenderAsync(CancellationToken cancellationToken) {
        if (_exportAsync == null) return Render(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        Options.ValidateImageExportOptions();
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Rendering, 0, 1));
        OfficeImageExportResult result = await _exportAsync(_format, Options, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        result.Require(Options.Policy);
        return result;
    }

    private TBuilder This => (TBuilder)this;
}
