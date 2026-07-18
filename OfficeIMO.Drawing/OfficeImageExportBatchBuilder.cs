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

    /// <summary>Configures a print profile with an explicit target DPI.</summary>
    public TBuilder ForPrint(double dpi = 300D) {
        if (dpi <= 0D || double.IsNaN(dpi) || double.IsInfinity(dpi)) throw new ArgumentOutOfRangeException(nameof(dpi));
        _format = OfficeImageExportFormat.Png;
        Options.TargetDpi = dpi;
        Options.BackgroundColor = OfficeColor.White;
        return This;
    }

    /// <summary>Exports all selected images and materializes their encoded payloads.</summary>
    public IReadOnlyList<OfficeImageExportResult> Export() {
        var results = new List<OfficeImageExportResult>();
        ExportEach(results.Add);
        return results.AsReadOnly();
    }

    /// <summary>Asynchronously exports all selected images and materializes their encoded payloads.</summary>
    public async Task<IReadOnlyList<OfficeImageExportResult>> ExportAsync(
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        await ExportEachAsync(
            (result, _) => {
                results.Add(result);
                return Task.CompletedTask;
            },
            cancellationToken).ConfigureAwait(false);
        return results.AsReadOnly();
    }

    /// <summary>Streams results to a consumer so callers can release each payload immediately.</summary>
    public void ExportEach(OfficeImageExportConsumer consumer, CancellationToken cancellationToken = default) {
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        cancellationToken.ThrowIfCancellationRequested();
        Options.ValidateImageExportOptions();
        var tracker = new OfficeImageExportBatchTracker(Options);
        int completed = 0;
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Rendering, 0));

        void Accept(OfficeImageExportResult result) {
            cancellationToken.ThrowIfCancellationRequested();
            result.Require(Options.Policy);
            tracker.Add(result);
            consumer(result);
            completed++;
            Options.Progress?.Report(new OfficeImageExportProgress(
                OfficeImageExportProgressStage.Completed,
                completed,
                name: result.Name));
        }

        if (_exportEach != null) {
            _exportEach(_format, Options, Accept, cancellationToken);
            return;
        }

        foreach (OfficeImageExportResult result in _export(_format, Options)) Accept(result);
    }

    /// <summary>Streams results asynchronously when the adapter owns a genuine asynchronous path.</summary>
    public async Task ExportEachAsync(
        OfficeImageExportAsyncConsumer consumer,
        CancellationToken cancellationToken = default) {
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        cancellationToken.ThrowIfCancellationRequested();
        Options.ValidateImageExportOptions();
        var tracker = new OfficeImageExportBatchTracker(Options);
        int completed = 0;
        Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Rendering, 0));

        async Task AcceptAsync(OfficeImageExportResult result, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            result.Require(Options.Policy);
            tracker.Add(result);
            await consumer(result, token).ConfigureAwait(false);
            completed++;
            Options.Progress?.Report(new OfficeImageExportProgress(
                OfficeImageExportProgressStage.Completed,
                completed,
                name: result.Name));
        }

        if (_exportEachAsync != null) {
            await _exportEachAsync(_format, Options, AcceptAsync, cancellationToken).ConfigureAwait(false);
            return;
        }

        if (_exportAsync != null) {
            IReadOnlyList<OfficeImageExportResult> asyncResults =
                await _exportAsync(_format, Options, cancellationToken).ConfigureAwait(false);
            foreach (OfficeImageExportResult result in asyncResults) await AcceptAsync(result, cancellationToken).ConfigureAwait(false);
            return;
        }

        if (_exportEach != null) {
            var pending = new List<OfficeImageExportResult>();
            _exportEach(_format, Options, pending.Add, cancellationToken);
            foreach (OfficeImageExportResult result in pending) await AcceptAsync(result, cancellationToken).ConfigureAwait(false);
            return;
        }

        foreach (OfficeImageExportResult result in _export(_format, Options)) {
            await AcceptAsync(result, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>Saves all selected images and returns payload-bearing results with their normalized paths.</summary>
    public IReadOnlyList<OfficeImageExportResult> Save(string folderPath) {
        string fullFolder = PrepareFolder(folderPath);
        var saved = new List<OfficeImageExportResult>();
        var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int index = 0;
        ExportEach(result => {
            string path = ResolveBatchPath(fullFolder, result, index++, usedFileNames);
            Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, index - 1, name: result.Name, destinationPath: path));
            OfficeFileCommit.WriteAllBytes(path, result.Bytes, OfficeImageExportPath.ToCommitPolicy(_conflictPolicy));
            saved.Add(result.WithSavedPath(path));
        });
        return saved.AsReadOnly();
    }

    /// <summary>
    /// Saves all selected images and returns only paths, metadata, and diagnostics so encoded payloads can be released.
    /// </summary>
    public OfficeImageExportBatchSaveResult SaveFiles(string folderPath) {
        string fullFolder = PrepareFolder(folderPath);
        var files = new List<OfficeImageExportSavedFile>();
        var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int index = 0;
        ExportEach(result => {
            string path = ResolveBatchPath(fullFolder, result, index++, usedFileNames);
            Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, index - 1, name: result.Name, destinationPath: path));
            OfficeFileCommit.WriteAllBytes(path, result.Bytes, OfficeImageExportPath.ToCommitPolicy(_conflictPolicy));
            files.Add(new OfficeImageExportSavedFile(result, path));
        });
        return new OfficeImageExportBatchSaveResult(files);
    }

    /// <summary>Asynchronously saves all selected images and returns payload-bearing results with normalized paths.</summary>
    public async Task<IReadOnlyList<OfficeImageExportResult>> SaveAsync(
        string folderPath,
        CancellationToken cancellationToken = default) {
        string fullFolder = PrepareFolder(folderPath);
        var saved = new List<OfficeImageExportResult>();
        var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int index = 0;
        await ExportEachAsync(async (result, token) => {
            string path = ResolveBatchPath(fullFolder, result, index++, usedFileNames);
            Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, index - 1, name: result.Name, destinationPath: path));
            await OfficeFileCommit.WriteAllBytesAsync(
                path,
                result.Bytes,
                OfficeImageExportPath.ToCommitPolicy(_conflictPolicy),
                token).ConfigureAwait(false);
            saved.Add(result.WithSavedPath(path));
        }, cancellationToken).ConfigureAwait(false);
        return saved.AsReadOnly();
    }

    /// <summary>Asynchronously saves images without retaining encoded payloads in the returned result.</summary>
    public async Task<OfficeImageExportBatchSaveResult> SaveFilesAsync(
        string folderPath,
        CancellationToken cancellationToken = default) {
        string fullFolder = PrepareFolder(folderPath);
        var files = new List<OfficeImageExportSavedFile>();
        var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int index = 0;
        await ExportEachAsync(async (result, token) => {
            string path = ResolveBatchPath(fullFolder, result, index++, usedFileNames);
            Options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, index - 1, name: result.Name, destinationPath: path));
            await OfficeFileCommit.WriteAllBytesAsync(
                path,
                result.Bytes,
                OfficeImageExportPath.ToCommitPolicy(_conflictPolicy),
                token).ConfigureAwait(false);
            files.Add(new OfficeImageExportSavedFile(result, path));
        }, cancellationToken).ConfigureAwait(false);
        return new OfficeImageExportBatchSaveResult(files);
    }

    private string ResolveBatchPath(
        string folder,
        OfficeImageExportResult result,
        int index,
        ISet<string> usedFileNames) {
        string name = string.IsNullOrWhiteSpace(result.Name)
            ? "image-" + (index + 1).ToString(CultureInfo.InvariantCulture)
            : result.Name!;
        string fileName = GetUniqueFileName(SanitizeFileName(name), _format.GetFileExtension(), usedFileNames);
        return OfficeImageExportPath.ResolveFile(Path.Combine(folder, fileName), _format, _conflictPolicy);
    }

    private static string PrepareFolder(string folderPath) {
        if (string.IsNullOrWhiteSpace(folderPath)) throw new ArgumentException("Output folder cannot be null or whitespace.", nameof(folderPath));
        string fullFolder = Path.GetFullPath(folderPath);
        Directory.CreateDirectory(fullFolder);
        return fullFolder;
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
            candidate.Equals("NUL", StringComparison.OrdinalIgnoreCase)) return true;
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
        if (string.IsNullOrWhiteSpace(baseName)) baseName = "image";
        string candidate = baseName + extension;
        if (usedFileNames.Add(candidate)) return candidate;

        int suffix = 2;
        do {
            candidate = baseName + "-" + suffix.ToString(CultureInfo.InvariantCulture) + extension;
            suffix++;
        } while (!usedFileNames.Add(candidate));
        return candidate;
    }

    private TBuilder This => (TBuilder)this;
}
