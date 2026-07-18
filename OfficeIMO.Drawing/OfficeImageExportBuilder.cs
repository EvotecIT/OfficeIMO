using OfficeIMO.Drawing.Internal;
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
    private readonly Func<OfficeImageExportFormat, TOptions, CancellationToken, Task<OfficeImageExportResult>>? _exportAsync;
    private OfficeImageExportFormat _format = OfficeImageExportFormat.Png;

    /// <summary>
    /// Creates a fluent export builder over an existing document-specific export function.
    /// </summary>
    protected OfficeImageExportBuilder(TOptions options, Func<OfficeImageExportFormat, TOptions, OfficeImageExportResult> export) {
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _export = export ?? throw new ArgumentNullException(nameof(export));
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

    /// <summary>Configures a standard preview profile: PNG, 1x scale, white background.</summary>
    public TBuilder ForPreview() {
        _format = OfficeImageExportFormat.Png;
        Options.Scale = 1D;
        Options.BackgroundColor = OfficeColor.White;
        return This;
    }

    /// <summary>Configures a high-resolution profile: PNG, 2x scale, white background.</summary>
    public TBuilder ForHighResolution() {
        _format = OfficeImageExportFormat.Png;
        Options.Scale = 2D;
        Options.BackgroundColor = OfficeColor.White;
        return This;
    }

    /// <summary>Exports using the currently configured format and options.</summary>
    public OfficeImageExportResult Export() => _export(_format, Options);

    /// <summary>Exports and returns the encoded image bytes.</summary>
    public byte[] ToBytes() => Export().Bytes;

    /// <summary>Saves the exported image to a file path.</summary>
    public OfficeImageExportResult Save(string path) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));
        }

        OfficeImageExportResult result = Export();
        OfficeFileCommit.WriteAllBytes(path, result.Bytes);
        return result;
    }

    /// <summary>Writes the exported image to a stream.</summary>
    public OfficeImageExportResult Save(Stream stream) {
        OfficeImageExportResult result = Export();
        OfficeStreamWriter.WriteAllBytes(stream, result.Bytes);
        return result;
    }

    /// <summary>Asynchronously saves the exported image to a file path.</summary>
    public async Task<OfficeImageExportResult> SaveAsync(
        string path,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));
        }

        cancellationToken.ThrowIfCancellationRequested();
        OfficeImageExportResult result = await ExportForSaveAsync(cancellationToken).ConfigureAwait(false);
        await OfficeFileCommit.WriteAllBytesAsync(path, result.Bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
        return result;
    }

    /// <summary>Asynchronously writes the exported image to a stream.</summary>
    public async Task<OfficeImageExportResult> SaveAsync(
        Stream stream,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeImageExportResult result = await ExportForSaveAsync(cancellationToken).ConfigureAwait(false);
        await OfficeStreamWriter.WriteAllBytesAsync(stream, result.Bytes, cancellationToken).ConfigureAwait(false);
        return result;
    }

    private async Task<OfficeImageExportResult> ExportForSaveAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_exportAsync != null) {
            return await _exportAsync(_format, Options, cancellationToken).ConfigureAwait(false);
        }
        OfficeImageExportResult result = _export(_format, Options);
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }

    private TBuilder This => (TBuilder)this;
}
