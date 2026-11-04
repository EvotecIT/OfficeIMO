using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Reader;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Reader;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Ocr;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Conversion;

public sealed partial class ConvertOfficeDocumentCommand {
    /// <summary>Recognize an original image or PDF before converting through the native OCR projections. OcrOptions also enables recognition.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter Ocr { get; set; }

    /// <summary>Optional provider, session and recognition policies. Captured before execution; never serialized into a job.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    [OfficeOptionsTransformation(typeof(OfficeOcrOptions))]
    public OfficeOcrOptions? OcrOptions { get; set; }

    /// <summary>One-based PDF source pages. An existing OCR review already owns its recognized selection.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    [OfficePageRangeTransformation]
    public OfficePageRange? Page { get; set; }

    /// <summary>Raster density for OCR and PDF to Word or PowerPoint, from 36 through 600.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "PathFile")]
    [Parameter(ParameterSetName = "LiteralPathFile")]
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Document")]
    [ValidateRange(36D, 600D)]
    public double? Dpi { get; set; }

    private OfficeOcrOptions? _ocrSettings;
    private Lazy<Task<OfficeIMO.Ocr.IOcrEngine>>? _ocrEngine;
    private bool HasOcrRequest => Ocr.IsPresent || OcrOptions != null;
    private bool UsesOcr => HasOcrRequest || ParameterSetName == "Document" && OfficeArgumentValue.Unwrap(Document) is PdfSearchableOcrReview;

    private void CaptureOcrSettings(string target) {
        OfficeImageExporter.ValidateExtension(target);
        if (Route != null || Theme != null || CompressPdf || AllowLegacyImportLoss || OutputProfile != OfficeConversionProfile.Faithful)
            throw new PSArgumentException("OCR uses native recognized-content projections; Route, Theme, CompressPdf, AllowLegacyImportLoss and alternate OutputProfile are not supported.");
        if (Plan || SaveJob != null || Resume || RetryFailed || CheckpointPath != null || ReportPath != null || JobName != null || PasswordVariable != null || PdfPasswordVariable != null)
            throw new PSArgumentException("OCR provider sessions and reviews are execution-only and cannot be used with portable jobs, plans, reports or checkpoints.");
        _ocrSettings ??= OfficeReadOcr.Capture(OcrOptions ?? new OfficeOcrOptions());
        _ocrEngine ??= new Lazy<Task<OfficeIMO.Ocr.IOcrEngine>>(() => OfficeReadOcr.CreateEngineAsync(_ocrSettings, CancelToken));
        if (Page != null) _ocrSettings.RecognitionOptions.ReadOptions.PageSelection = Page.Selection;
        if (Dpi.HasValue) _ocrSettings.RecognitionOptions.Dpi = Dpi.Value;
    }

    private void ApplyPageAndDpi(OfficeConversionOptions options, System.Collections.Generic.IReadOnlyList<OfficeConversionCapability> routes) {
        if (UsesOcr || Page == null && !Dpi.HasValue) return;
        if (routes.Count == 0 || routes.Any(route => !route.SourceExtensions.Contains(".pdf", StringComparer.OrdinalIgnoreCase)))
            throw new PSArgumentException("Page and Dpi require a supported PDF source route, or OCR.");
        if (Page != null) (options.PdfLoad ??= new PdfLoadOptions()).PageSelection = Page.Selection;
        if (Dpi.HasValue) {
            foreach (var route in routes) {
                if (route.TargetExtension == ".docx") (options.PdfWord ??= new OfficeIMO.Pdf.Word.PdfWordExportOptions()).Dpi = Dpi.Value;
                else if (route.TargetExtension == ".pptx") (options.PdfPowerPoint ??= new OfficeIMO.Pdf.PowerPoint.PdfPowerPointExportOptions()).Dpi = Dpi.Value;
                else throw new PSArgumentException("Dpi applies to OCR and PDF to Word or PowerPoint raster projections.");
            }
        }
    }

    private async Task<PdfSearchableOcrReview> RecognizePathAsync(string input, OfficeConversionOptions options, CancellationToken token) {
        var captured = options.Clone();
        var recognition = OfficeReadOcr.Capture(_ocrSettings!);
        var reader = captured.Reader ?? new ReaderOptions();
        reader.MaxInputBytes = Math.Min(reader.MaxInputBytes ?? long.MaxValue, captured.MaximumInputBytes);
        if (captured.SourcePassword != null) reader.OpenPassword = captured.SourcePassword;
        if (captured.SourceFormat != null || captured.SourceExtension != null)
            throw new PSArgumentException("OCR admits original image or PDF files by their native source contract; omit From and source-codec overrides.");
        var load = (captured.PdfLoad ?? recognition.RecognitionOptions.ReadOptions).Clone(reader.MaxInputBytes, reader.OpenPassword);
        if (Page != null) load.PageSelection = Page.Selection;
        return await OfficeReadOcr.ReadReviewAsync(input, _ => _ocrEngine!.Value,
            recognition, reader, load, token).ConfigureAwait(false);
    }

    private async Task<OfficeImageConversionResult> ConvertOcrPathAsync(string input, string output, OfficeConversionOptions options, CancellationToken token) {
        if (OfficePath.AreEquivalent(input, output)) throw new PSArgumentException("The conversion destination must differ from the original OCR source.");
        OfficeBatchExecutor.EnsureNoLinks(input);
        OfficeBatchExecutor.EnsureNoLinks(output);
        var review = await RecognizePathAsync(input, options, token).ConfigureAwait(false);
        return await OfficeImageExporter.ExportAsync(review, output, options, token).ConfigureAwait(false);
    }

    private async Task ConvertOcrReviewAsync(PdfSearchableOcrReview review) {
        if (Page != null || Dpi.HasValue || OcrOptions != null || Password != null || From != null)
            throw new PSArgumentException("An OCR review already captures its source and recognition settings. Convert it directly without Page, Dpi, OcrOptions, Password or From.");
        var target = OfficeFormatSelection.Parse(To ?? System.IO.Path.GetExtension(OutputPath))!;
        var options = CaptureOptions(new[] { "review.pdf" }, target.Extension);
        if (AsString && OutputPath != null) throw new PSArgumentException("AsString and OutputPath cannot be combined.");
        if (OutputPath == null) {
            OfficeImageExporter.ValidateTextExtension(target.Extension);
            if (!ShouldProcess("in-memory " + target.Extension + " output", "Convert OCR review")) return;
            string text = OfficeImageExporter.ExportText(review, target.Extension, options, CancelToken, out var report);
            SetSummary(report);
            WriteObject(text);
            return;
        }
        string output = OfficePath.ResolveOutput(SessionState, OutputPath, directory: false);
        if (!target.Matches(System.IO.Path.GetExtension(output))) throw new PSArgumentException("To and the OutputPath extension must name the same target format.");
        if (!ShouldProcess(output, "Convert OCR review")) return;
        var result = await OfficeImageExporter.ExportAsync(review, output, options, CancelToken);
        SetSummary(result);
        WriteObject(result);
        OpenPublishedOutput(result);
    }
}
