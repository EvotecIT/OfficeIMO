using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Shared native image orchestration with provider admission before background work or publication.</summary>
public abstract class OfficeDocumentImageCmdlet : AsyncPSCmdlet {
    private readonly List<Source> _sources = new();
    /// <summary>Gets or sets wildcard input paths.</summary>
    public abstract string[] Path { get; set; }
    /// <summary>Gets or sets literal input paths.</summary>
    public abstract string[] LiteralPath { get; set; }
    /// <summary>Gets or sets a caller-owned native document.</summary>
    public abstract object Document { get; set; }
    /// <summary>Gets or sets literal markup input.</summary>
    public abstract string Text { get; set; }
    /// <summary>Gets or sets the markup source format.</summary>
    public abstract string From { get; set; }
    /// <summary>Gets whether this command publishes image files.</summary>
    protected abstract bool WritesFiles { get; }
    /// <summary>Gets the exact single-image destination, when selected.</summary>
    protected virtual string? ImageOutputPath => null;
    /// <summary>Gets the destination directory for a batch.</summary>
    protected virtual string? ImageOutputDirectory => null;
    /// <summary>Gets whether an existing image must remain unchanged.</summary>
    protected virtual bool PreventOverwrite => false;

    /// <summary>Output image encoding. File destinations infer it when omitted.</summary>
    [Parameter]
    public OfficeImageFormat Format { get; set; } = OfficeImageFormat.Png;
    /// <summary>One-based caller-ordered pages, slides or diagram pages; defaults to all.</summary>
    [Parameter]
    [OfficePageRangeTransformation]
    public OfficePageRange? Page { get; set; }
    /// <summary>Excel worksheets selected by name, one-based position or native sheet object.</summary>
    [Parameter]
    [OfficeSheetTransformation]
    public OfficeSheetSelector[]? Sheet { get; set; }
    /// <summary>A1 range within the selected Excel sheets.</summary>
    [Parameter]
    public string? Range { get; set; }
    /// <summary>Chart name within the selected Excel sheets. Cannot be combined with Range.</summary>
    [Parameter]
    public string? Chart { get; set; }
    /// <summary>Target image resolution in dots per inch; omitted values retain native options.</summary>
    [Parameter]
    public double Dpi { get; set; }
    /// <summary>Image scale; omitted values retain native options.</summary>
    [Parameter]
    public double Scale { get; set; } = 1D;
    /// <summary>Image background, including Transparent.</summary>
    [Parameter]
    [OfficeColorTransformation]
    public OfficeColor BackgroundColor { get; set; }
    /// <summary>Native image/load options or hashtables, routed to the selected source formats.</summary>
    [Parameter]
    public object[]? Options { get; set; }
    /// <summary>Password used only when opening encrypted source files.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>Maximum number of input files processed concurrently.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(1, 32)]
    public int ThrottleLimit { get; set; } = 4;

    /// <inheritdoc />
    protected override Task ProcessRecordAsync() {
        if (ParameterSetName.StartsWith("Text", StringComparison.Ordinal)) {
            var from = OfficeFormatSelection.Parse(From)!;
            if (!(from.Family == OfficeDocumentFormat.Html || from.Matches(".html")))
                throw new PSArgumentException("Text image input requires From Html.");
            _sources.Add(new Source(null, null, Text, OfficeDocumentFormat.Html));
        } else if (ParameterSetName.StartsWith("Document", StringComparison.Ordinal)) {
            if (OfficePath.GetPipelineLiteralPaths(Document) is { } paths) AddPaths(paths, literal: true);
            else {
                object native = OfficeArgumentValue.Unwrap(Document)!;
                _sources.Add(new Source(null, native, null, OfficeDocumentImageSource.Format(native)));
            }
        } else AddPaths(ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal) ? LiteralPath : Path,
            ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal));
        return Task.CompletedTask;
    }

    private void AddPaths(IEnumerable<string> inputs, bool literal) {
        foreach (string input in inputs) {
            try {
                foreach (string path in OfficePath.ResolveInput(SessionState, input, literal, allowDirectory: false,
                    (error, item) => WriteError(OfficeError.Create(error, "input-unreadable", item))))
                    _sources.Add(new Source(path, null, null, OfficeDocumentImageSource.Format(path)));
            } catch (ArgumentException error) {
                ThrowTerminatingError(OfficeError.Create(error, "invalid-argument", input, ErrorCategory.InvalidArgument));
            } catch (Exception error) when ((error is IOException or RuntimeException or UnauthorizedAccessException)
                && error is not (ActionPreferenceStopException or PipelineStoppedException)) {
                WriteError(OfficeError.Create(error, "input-unreadable", input));
            }
        }
    }

    /// <inheritdoc />
    protected override async Task EndProcessingAsync() {
        if (_sources.Count == 0) return;
        bool single = ImageOutputPath != null;
        if (single && _sources.Count != 1) throw new PSArgumentException("OutputPath requires exactly one input and one image; use OutputDirectory for a batch.");
        if (Range != null && Chart != null) throw new PSArgumentException("Range and Chart cannot be combined.");
        var formats = _sources.Select(source => source.Format).Distinct().ToArray();
        if ((Sheet != null || Range != null || Chart != null) && formats.Any(format => format != OfficeDocumentFormat.Excel))
            throw new PSArgumentException("Sheet, Range and Chart require Excel sources.");
        if (Page != null && formats.Contains(OfficeDocumentFormat.Excel)) throw new PSArgumentException("Use Sheet to select Excel worksheets.");
        bool hasPassword = MyInvocation.BoundParameters.ContainsKey(nameof(Password));
        if (hasPassword && (_sources.Any(source => source.Document != null || source.Text != null)
            || formats.Any(format => format is OfficeDocumentFormat.Html or OfficeDocumentFormat.Visio)))
            throw new PSArgumentException("Password applies only to encrypted Word, Excel, PowerPoint or PDF input files.");
        string? outputPath = ImageOutputPath == null ? null : OfficePath.ResolveOutput(SessionState, ImageOutputPath, directory: false);
        string? outputDirectory = ImageOutputDirectory == null ? null : OfficePath.ResolveOutput(SessionState, ImageOutputDirectory, directory: true);
        OfficeImageExportFormat encoding = ResolveEncoding(outputPath);
        if (outputPath != null) outputPath = OfficeImageExportPath.NormalizeFile(outputPath, encoding);
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        OfficeDocumentImageOptions options = OfficeDocumentImageOptions.Capture(engine, Options, formats);
        if (formats.Contains(OfficeDocumentFormat.Visio) && !string.IsNullOrWhiteSpace(options.Visio.FontFilePath))
            options.Visio.FontFilePath = OfficePath.ResolveOutput(SessionState, options.Visio.FontFilePath!, directory: false);
        foreach (OfficeDocumentFormat format in formats) {
            OfficeImageExportOptions image = options.Image(format);
            image.Format = encoding;
            if (MyInvocation.BoundParameters.ContainsKey(nameof(Dpi))) image.TargetDpi = Dpi;
            if (MyInvocation.BoundParameters.ContainsKey(nameof(Scale))) image.Scale = Scale;
            if (MyInvocation.BoundParameters.ContainsKey(nameof(BackgroundColor))) image.BackgroundColor = BackgroundColor;
            if (single) image.MaximumOutputCount = Math.Min(image.MaximumOutputCount, 1);
            image.ValidateImageExportOptions();
            if (format == OfficeDocumentFormat.Html && encoding is not (OfficeImageExportFormat.Png or OfficeImageExportFormat.Jpeg
                or OfficeImageExportFormat.Svg or OfficeImageExportFormat.Tiff or OfficeImageExportFormat.Webp))
                throw new PSArgumentException("The native HTML exporter does not support this image encoding.");
        }
        if (hasPassword && Password != null) {
            string password = new NetworkCredential(string.Empty, Password).Password;
            options.WordLoad.Password = password; options.ExcelLoad.Password = password;
            options.PowerPointLoad.Password = password; options.PdfLoad = options.PdfLoad.Clone(maximumInputBytes: null, passwordOverride: password);
        }
        var request = new RenderRequest(encoding, Page, Sheet?.ToArray(), Range, Chart, outputPath, outputDirectory, WritesFiles, PreventOverwrite);
        if (request.Save && !ShouldProcess(outputPath ?? outputDirectory, "Export document images")) return;
        if (_sources.All(source => source.Path != null)) {
            await OfficeBatch.RunAsync(_sources, async (source, token) => {
                await RenderAsync(source, options.Clone(), request, token).ConfigureAwait(false);
                return null;
            }, single ? 1 : ThrottleLimit, CancelToken, this, "Render document images", "image-export-failed");
        } else {
            // Caller-owned native objects are never disposed or rendered concurrently.
            foreach (Source source in _sources) {
                CancelToken.ThrowIfCancellationRequested();
                await RenderAsync(source, options.Clone(), request, CancelToken);
            }
        }
    }

    private OfficeImageExportFormat ResolveEncoding(string? outputPath) {
        OfficeImageExportFormat[] supported = (OfficeImageExportFormat[])Enum.GetValues(typeof(OfficeImageExportFormat));
        OfficeImageExportFormat? inferred = outputPath == null ? null
            : supported.Where(format => format.HasFileExtension(System.IO.Path.GetExtension(outputPath))).Select(format => (OfficeImageExportFormat?)format).SingleOrDefault();
        if (outputPath != null && !string.IsNullOrEmpty(System.IO.Path.GetExtension(outputPath)) && inferred == null) throw new PSArgumentException("OutputPath must have a supported image extension.");
        OfficeImageExportFormat? selected = supported.Where(format => format.GetContainerFormat() == Format).Select(format => (OfficeImageExportFormat?)format).SingleOrDefault();
        if (selected == null) throw new PSArgumentException("The selected image format has no native encoder.");
        if (inferred.HasValue && MyInvocation.BoundParameters.ContainsKey(nameof(Format)) && inferred != selected)
            throw new PSArgumentException("Format and OutputPath extension must match.");
        return inferred ?? selected.Value;
    }

    private async Task RenderAsync(Source source, OfficeDocumentImageOptions options, RenderRequest request, CancellationToken token) {
        object? owned = null;
        try {
            object document = source.Document ?? (source.Path != null
                ? owned = await OfficeDocumentImageSource.LoadAsync(source.Path, source.Format, options, token).ConfigureAwait(false)
                : HtmlConversionDocument.Parse(source.Text!, options.HtmlLoad, token));
            IReadOnlyList<OfficeImageExportResult> images = OfficeDocumentImageRenderer.Render(document, request.Encoding,
                options, request.Pages, request.Sheets, request.Range, request.Chart, token);
            if (request.OutputPath != null && images.Count != 1) throw new PSArgumentException("OutputPath requires exactly one image; use Page or OutputDirectory.");
            string baseName = OfficePortableFileName.SanitizeBaseName(System.IO.Path.GetFileNameWithoutExtension(
                source.Path ?? (source.Document == null ? "source.html" : OfficeLiveDocumentSource.Create(source.Document).Path ?? "source")));
            for (int index = 0; index < images.Count; index++) {
                token.ThrowIfCancellationRequested();
                OfficeImageExportResult image = images[index];
                if (request.Save) {
                    string suffix = source.Format == OfficeDocumentFormat.Excel && (request.Range != null || request.Chart != null)
                        ? OfficePortableFileName.SanitizeBaseName((image.Source ?? image.Name ?? (index + 1).ToString()).Replace('!', '-'))
                        : (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string path = request.OutputPath ?? System.IO.Path.Combine(request.OutputDirectory!, baseName + "-" + suffix + request.Encoding.GetFileExtension());
                    image = image.Save(path, request.NoClobber ? OfficeImageExportFileConflictPolicy.FailIfExists : OfficeImageExportFileConflictPolicy.Replace);
                }
                WriteObject(image);
            }
        } finally { (owned as IDisposable)?.Dispose(); }
    }

    private sealed class RenderRequest {
        internal RenderRequest(OfficeImageExportFormat encoding, OfficePageRange? pages, OfficeSheetSelector[]? sheets,
            string? range, string? chart, string? outputPath, string? outputDirectory, bool save, bool noClobber) {
            Encoding = encoding; Pages = pages; Sheets = sheets; Range = range; Chart = chart;
            OutputPath = outputPath; OutputDirectory = outputDirectory; Save = save; NoClobber = noClobber;
        }
        internal OfficeImageExportFormat Encoding { get; }
        internal OfficePageRange? Pages { get; }
        internal OfficeSheetSelector[]? Sheets { get; }
        internal string? Range { get; }
        internal string? Chart { get; }
        internal string? OutputPath { get; }
        internal string? OutputDirectory { get; }
        internal bool Save { get; }
        internal bool NoClobber { get; }
    }

    private sealed class Source {
        internal Source(string? path, object? document, string? text, OfficeDocumentFormat format) { Path = path; Document = document; Text = text; Format = format; }
        internal string? Path { get; }
        internal object? Document { get; }
        internal string? Text { get; }
        internal OfficeDocumentFormat Format { get; }
        public override string ToString() => Path ?? Format + (Document == null ? " text" : " document");
    }
}
