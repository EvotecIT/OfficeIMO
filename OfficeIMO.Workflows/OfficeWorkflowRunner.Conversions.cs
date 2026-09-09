using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static OperationArtifact Convert(
        ValidatedRequest request,
        List<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        byte[] input = OfficeWorkflowInputReader.ReadAllBytes(
            request.InputPath,
            request.Limits.MaximumInputBytes,
            cancellationToken);
        return Convert(request, input, diagnostics, cancellationToken);
    }

    private static OperationArtifact Convert(
        ValidatedRequest request,
        byte[] input,
        List<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken,
        bool emitHtmlTaggedStructure = true,
        IReadOnlyDictionary<string, byte[]>? htmlResourceSnapshots = null) {
        ArgumentNullException.ThrowIfNull(input);
        OfficeWorkflowRoute route = request.Route!;
        OfficeWorkflowConversionOptions settings = request.ConversionOptions ?? new();
        PdfReadOptions? readOptions = settings.CreateReadOptions();
        cancellationToken.ThrowIfCancellationRequested();
        long maximumOutputBytes = request.Limits.MaximumOutputBytes;
        byte[] bytes;
        bool hasLoss = false;
        switch (route.Id) {
            case "docx-pdf":
                using (var source = new MemoryStream(input, writable: false))
                using (WordDocument document = WordDocument.LoadAsync(
                    source,
                    cancellationToken: cancellationToken).GetAwaiter().GetResult()) {
                    var options = new WordToPdfOptions();
                    options.UseProfile(ToPdfExportProfile(request.OutputProfile));
                    PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options, cancellationToken);
                    bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                    hasLoss = conversion.HasLoss;
                    AddPdfWarnings(conversion.Warnings, diagnostics);
                }
                break;
            case "xlsx-pdf":
                using (var source = new MemoryStream(input, writable: false))
                using (ExcelDocument document = ExcelDocument.LoadAsync(
                    source,
                    cancellationToken: cancellationToken).GetAwaiter().GetResult()) {
                    var options = new ExcelToPdfOptions();
                    options.UseProfile(ToPdfExportProfile(request.OutputProfile));
                    if (settings.WorksheetLayout.HasValue) options.WorksheetLayout = settings.WorksheetLayout.Value;
                    PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options, cancellationToken);
                    bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                    hasLoss = conversion.HasLoss;
                    AddPdfWarnings(conversion.Warnings, diagnostics);
                }
                break;
            case "pptx-pdf":
                using (var source = new MemoryStream(input, writable: false))
                using (PowerPointPresentation document = PowerPointPresentation.LoadAsync(
                    source,
                    cancellationToken: cancellationToken).GetAwaiter().GetResult()) {
                    var options = new PowerPointToPdfOptions();
                    options.UseProfile(ToPdfExportProfile(request.OutputProfile));
                    PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options, cancellationToken);
                    bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                    hasLoss = conversion.HasLoss;
                    AddPdfWarnings(conversion.Warnings, diagnostics);
                }
                break;
            case "html-pdf": {
                long remainingInputBytes = Math.Max(0L, request.Limits.MaximumInputBytes - input.LongLength);
                HtmlToPdfOptions options = OfficeWorkflowHtmlResourceResolver.CreateOptions(
                    request.InputPath,
                    remainingInputBytes,
                    htmlResourceSnapshots ?? (request.InputStream is null ? null : new Dictionary<string, byte[]>()));
                if (!emitHtmlTaggedStructure) {
                    options.PdfOptions.SetTaggedStructureMode(PdfTaggedStructureMode.None);
                }
                PdfDocumentConversionResult conversion = ParseHtmlInput(input, request.InputPath, cancellationToken)
                    .ToPdfDocumentResultAsync(options, cancellationToken)
                    .GetAwaiter()
                    .GetResult();
                bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss;
                AddPdfWarnings(conversion.Warnings, diagnostics);
                break;
            }
            case "pdf-docx": {
                PdfDocument pdf = PdfDocument.Load(input, request.PdfLoadOptions);
                PdfWordConversionResult conversion = pdf.ToWordDocumentResult(new PdfToWordOptions {
                    Mode = settings.WordMode ?? PdfWordImportMode.EditableContent,
                    ReadOptions = readOptions, Dpi = settings.RasterDpi ?? 144,
                    MaxTotalOutputBytes = maximumOutputBytes, MaxOutputBytesPerPage = Math.Min(64L * 1024L * 1024L, maximumOutputBytes)
                }, cancellationToken);
                using WordDocument document = conversion.Value;
                using (var stream = new OfficeWorkflowBoundedMemoryStream(maximumOutputBytes)) {
                    document.SaveAsync(stream, cancellationToken).GetAwaiter().GetResult();
                    bytes = stream.ToArray();
                }
                hasLoss = conversion.HasLoss;
                AddMessages(conversion.Report.Warnings.Select(static warning => warning.ToString()), hasLoss, diagnostics);
                break;
            }
            case "pdf-xlsx": {
                PdfDocument pdf = PdfDocument.Load(input, request.PdfLoadOptions);
                PdfExcelTableImportResult conversion = pdf.ImportTablesToExcelDocumentResult(new PdfTablesToExcelOptions { ReadOptions = readOptions }, cancellationToken);
                using ExcelDocument document = conversion.Value;
                using (var stream = new OfficeWorkflowBoundedMemoryStream(maximumOutputBytes)) {
                    document.SaveAsync(stream, cancellationToken).GetAwaiter().GetResult();
                    bytes = stream.ToArray();
                }
                hasLoss = conversion.HasLoss || conversion.HasOmittedPageContent;
                if (conversion.HasOmittedPageContent) {
                    diagnostics.Add(new OfficeWorkflowDiagnostic(
                        "PdfTablesOnly",
                        "Excel conversion reconstructs detected tables; other fixed-layout page content is outside this route.",
                        OfficeWorkflowDiagnosticSeverity.Warning,
                        "convert"));
                }
                break;
            }
            case "pdf-pptx": {
                PdfDocument pdf = PdfDocument.Load(input, request.PdfLoadOptions);
                PdfPowerPointConversionResult conversion = pdf.ToPowerPointPresentationResult(
                    new PdfToPowerPointOptions {
                        Mode = settings.PowerPointMode ?? PdfPowerPointImportMode.EditableContent,
                        ReadOptions = readOptions, Dpi = settings.RasterDpi ?? 144,
                        MaxTotalOutputBytes = maximumOutputBytes, MaxOutputBytesPerPage = Math.Min(64L * 1024L * 1024L, maximumOutputBytes)
                    }, cancellationToken);
                using PowerPointPresentation document = conversion.Value;
                using (var stream = new OfficeWorkflowBoundedMemoryStream(maximumOutputBytes)) {
                    document.SaveAsync(stream, cancellationToken).GetAwaiter().GetResult();
                    bytes = stream.ToArray();
                }
                hasLoss = conversion.HasLoss || conversion.HasOmittedPageContent;
                AddMessages(conversion.Warnings.Select(static warning => warning.ToString()), hasLoss, diagnostics);
                break;
            }
            case "pdf-html": {
                PdfDocument pdf = PdfDocument.Load(input, request.PdfLoadOptions);
                int maximumOutputCharacters = (int)Math.Min(int.MaxValue, maximumOutputBytes);
                PdfHtmlConversionResult conversion = pdf.ToHtmlResult(new PdfToHtmlOptions {
                    Profile = settings.HtmlProfile ?? PdfHtmlProfile.PositionedReview,
                    ReadOptions = readOptions,
                    IncludeLinkAnnotations = true,
                    IncludeFormWidgets = true,
                    MaximumOutputCharacters = maximumOutputCharacters,
                    MaxEmbeddedImageBytes = Math.Min(10L * 1024L * 1024L, maximumOutputBytes - maximumOutputBytes / 4L),
                }, cancellationToken);
                bytes = EncodeUtf8Bounded(conversion.Value, maximumOutputBytes);
                hasLoss = conversion.HasLoss;
                AddMessages(conversion.Report.Warnings.Select(static warning => warning.ToString()), hasLoss, diagnostics);
                break;
            }
            default:
                throw new NotSupportedException("The conversion route '" + route.Id + "' is not implemented by the local runner.");
        }

        if (settings.CompressPdfOutput) {
            PdfOptimizationOptions compression = PdfOptimizationOptions.Create(PdfOptimizationProfile.MaximumCompression);
            compression.CancellationToken = cancellationToken;
            compression.MaximumOutputBytes = maximumOutputBytes;
            PdfOptimizationActionResult optimized = PdfDocument.Load(bytes).Optimization.Apply(compression);
            if (!optimized.PreservationReport.IsPreserved) throw new InvalidOperationException("PDF compression did not preserve the converted document.");
            bytes = optimized.Bytes;
            diagnostics.Add(new OfficeWorkflowDiagnostic("PdfOutputCompression", "Verified lossless PDF compression completed; saved " + optimized.SavedBytes + " bytes.",
                OfficeWorkflowDiagnosticSeverity.Information, "convert"));
        }

        diagnostics.Add(new OfficeWorkflowDiagnostic(
            "RouteContract",
            route.Description,
            OfficeWorkflowDiagnosticSeverity.Information,
            "convert",
            new Dictionary<string, string>(StringComparer.Ordinal) {
                ["route"] = route.Id,
                ["engine"] = route.Engine,
                ["fidelity"] = route.Fidelity,
                ["supportLevel"] = route.SupportLevel,
                ["knownLimitations"] = route.KnownLimitations
            }));
        cancellationToken.ThrowIfCancellationRequested();
        string summary = hasLoss
            ? route.Label + " completed with fidelity warnings; review the structured diagnostics."
            : route.Label + " completed and the output reopened successfully.";
        return new OperationArtifact(bytes, summary, null);
    }

    private static string DecodeHtmlInput(byte[] input, CancellationToken cancellationToken) {
        using var source = new MemoryStream(input, writable: false);
        return HtmlConversionDocument.LoadAsync(
            source,
            cancellationToken: cancellationToken).GetAwaiter().GetResult().SourceHtml;
    }

    internal static HtmlConversionDocument ParseHtmlInput(
        byte[] input,
        string inputPath,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(input);
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("Input path cannot be empty.", nameof(inputPath));
        using var source = new MemoryStream(input, writable: false);
        return HtmlConversionDocument.LoadAsync(
            source,
            new HtmlConversionDocumentOptions {
                BaseUri = new Uri(Path.GetFullPath(inputPath)),
                ResourceUrlPolicy = OfficeWorkflowHtmlResourceResolver.CreateResourcePolicy()
            },
            cancellationToken: cancellationToken).GetAwaiter().GetResult();
    }
}
