using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared fluent image-export surface for document-level batch exports.
/// </summary>
/// <typeparam name="TBuilder">Concrete builder type returned from fluent methods.</typeparam>
/// <typeparam name="TOptions">Document-specific image export options.</typeparam>
public abstract class OfficeImageExportBatchBuilder<TBuilder, TOptions>
    where TBuilder : OfficeImageExportBatchBuilder<TBuilder, TOptions>
    where TOptions : OfficeImageExportOptions {
    private const int MaximumPortableBaseNameLength = 120;
    private const string PortableInvalidFileNameCharacters = "<>:\"/\\|?*";
    private static readonly char[] PlatformInvalidFileNameCharacters = Path.GetInvalidFileNameChars();
    private readonly Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> _export;
    private readonly Func<OfficeImageExportFormat, TOptions, CancellationToken, Task<IReadOnlyList<OfficeImageExportResult>>>? _exportAsync;
    private OfficeImageExportFormat _format = OfficeImageExportFormat.Png;

    /// <summary>
    /// Creates a fluent batch export builder over an existing document-specific export function.
    /// </summary>
    protected OfficeImageExportBatchBuilder(TOptions options, Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> export) {
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _export = export ?? throw new ArgumentNullException(nameof(export));
    }

    /// <summary>
    /// Creates a fluent batch builder with a genuine asynchronous renderer for resource-aware document models.
    /// </summary>
    protected OfficeImageExportBatchBuilder(
        TOptions options,
        Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> export,
        Func<OfficeImageExportFormat, TOptions, CancellationToken, Task<IReadOnlyList<OfficeImageExportResult>>> exportAsync)
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

    /// <summary>Sets the maximum pixel allocation for each raster result.</summary>
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

    /// <summary>Exports all selected images using the currently configured format and options.</summary>
    public IReadOnlyList<OfficeImageExportResult> Export() => _export(_format, Options);

    /// <summary>Saves all selected images to a folder.</summary>
    public IReadOnlyList<OfficeImageExportResult> Save(string folderPath) {
        if (string.IsNullOrWhiteSpace(folderPath)) {
            throw new ArgumentException("Output folder cannot be null or whitespace.", nameof(folderPath));
        }

        string fullFolder = Path.GetFullPath(folderPath);
        Directory.CreateDirectory(fullFolder);
        IReadOnlyList<OfficeImageExportResult> results = Export();
        string extension = _format.GetFileExtension();
        var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < results.Count; i++) {
            OfficeImageExportResult result = results[i];
            string name = string.IsNullOrWhiteSpace(result.Name)
                ? "image-" + (i + 1).ToString(CultureInfo.InvariantCulture)
                : result.Name!;
            string fileName = GetUniqueFileName(SanitizeFileName(name), extension, usedFileNames);
            OfficeFileCommit.WriteAllBytes(Path.Combine(fullFolder, fileName), result.Bytes);
        }

        return results;
    }

    private async Task<IReadOnlyList<OfficeImageExportResult>> ExportForSaveAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_exportAsync != null) {
            return await _exportAsync(_format, Options, cancellationToken).ConfigureAwait(false);
        }
        IReadOnlyList<OfficeImageExportResult> results = _export(_format, Options);
        cancellationToken.ThrowIfCancellationRequested();
        return results;
    }

    /// <summary>Asynchronously saves all selected images to a folder.</summary>
    public async Task<IReadOnlyList<OfficeImageExportResult>> SaveAsync(
        string folderPath,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(folderPath)) {
            throw new ArgumentException("Output folder cannot be null or whitespace.", nameof(folderPath));
        }

        cancellationToken.ThrowIfCancellationRequested();
        string fullFolder = Path.GetFullPath(folderPath);
        Directory.CreateDirectory(fullFolder);
        IReadOnlyList<OfficeImageExportResult> results = await ExportForSaveAsync(cancellationToken).ConfigureAwait(false);
        string extension = _format.GetFileExtension();
        var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < results.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = results[i];
            string name = string.IsNullOrWhiteSpace(result.Name)
                ? "image-" + (i + 1).ToString(CultureInfo.InvariantCulture)
                : result.Name!;
            string fileName = GetUniqueFileName(SanitizeFileName(name), extension, usedFileNames);
            await OfficeFileCommit.WriteAllBytesAsync(
                Path.Combine(fullFolder, fileName),
                result.Bytes,
                cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        return results;
    }

    private static string SanitizeFileName(string name) {
        char[] chars = name.ToCharArray();
        for (int i = 0; i < chars.Length; i++) {
            if (chars[i] < 32 ||
                PortableInvalidFileNameCharacters.IndexOf(chars[i]) >= 0 ||
                Array.IndexOf(PlatformInvalidFileNameCharacters, chars[i]) >= 0) {
                chars[i] = '_';
            }
        }

        string sanitized = new string(chars).Trim().TrimEnd('.', ' ');
        if (sanitized.Length > MaximumPortableBaseNameLength) {
            int length = MaximumPortableBaseNameLength;
            if (length > 0 && char.IsHighSurrogate(sanitized[length - 1])) length--;
            sanitized = sanitized.Substring(0, length).TrimEnd('.', ' ');
        }
        if (IsReservedWindowsFileName(sanitized)) sanitized = "_" + sanitized;
        return sanitized;
    }

    private static bool IsReservedWindowsFileName(string name) {
        if (string.IsNullOrWhiteSpace(name)) return false;
        string candidate = name;
        int dot = candidate.IndexOf('.');
        if (dot >= 0) candidate = candidate.Substring(0, dot);
        if (candidate.Equals("CON", StringComparison.OrdinalIgnoreCase) ||
            candidate.Equals("PRN", StringComparison.OrdinalIgnoreCase) ||
            candidate.Equals("AUX", StringComparison.OrdinalIgnoreCase) ||
            candidate.Equals("NUL", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }
        if (candidate.Length == 4 &&
            (candidate.StartsWith("COM", StringComparison.OrdinalIgnoreCase) ||
             candidate.StartsWith("LPT", StringComparison.OrdinalIgnoreCase))) {
            return IsReservedDeviceDigit(candidate[3]);
        }
        return false;
    }

    private static bool IsReservedDeviceDigit(char value) =>
        value >= '1' && value <= '9' ||
        value == '\u00B9' ||
        value == '\u00B2' ||
        value == '\u00B3';

    private static string GetUniqueFileName(string baseName, string extension, ISet<string> usedFileNames) {
        if (string.IsNullOrWhiteSpace(baseName)) {
            baseName = "image";
        }

        string candidate = baseName + extension;
        if (usedFileNames.Add(candidate)) {
            return candidate;
        }

        int suffix = 2;
        do {
            candidate = baseName + "-" + suffix.ToString(CultureInfo.InvariantCulture) + extension;
            suffix++;
        } while (!usedFileNames.Add(candidate));

        return candidate;
    }

    private TBuilder This => (TBuilder)this;
}
