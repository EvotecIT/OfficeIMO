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
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private const long MaximumOpenXmlPartCharacters = 10L * 1024L * 1024L;

    private static OfficePackageSecurityOptions CreateOpenXmlPackageSecurity(OfficeWorkflowLimits limits) {
        OfficePackageSecurityOptions security = OfficePackageSecurityOptions.SecureDefaults;
        security.MaxPackageBytes = limits.MaximumInputBytes;
        security.MaxXmlCharactersInPart = MaximumOpenXmlPartCharacters;
        return security;
    }

    private static OfficeOpenXmlLoadSettings CreateOpenXmlLoadSettings() => new() {
        MaxCharactersInPart = MaximumOpenXmlPartCharacters
    };

    private static WordLoadOptions CreateWordLoadOptions(OfficeWorkflowLimits limits) => new() {
        AccessMode = DocumentAccessMode.ReadOnly,
        MaxInputBytes = limits.MaximumInputBytes,
        PackageSecurity = CreateOpenXmlPackageSecurity(limits),
        OpenSettings = CreateOpenXmlLoadSettings()
    };

    private static ExcelLoadOptions CreateExcelLoadOptions(OfficeWorkflowLimits limits) => new() {
        AccessMode = DocumentAccessMode.ReadOnly,
        MaxInputBytes = limits.MaximumInputBytes,
        PackageSecurity = CreateOpenXmlPackageSecurity(limits),
        OpenSettings = CreateOpenXmlLoadSettings()
    };

    private static PowerPointLoadOptions CreatePowerPointLoadOptions(OfficeWorkflowLimits limits) => new() {
        AccessMode = DocumentAccessMode.ReadOnly,
        MaxInputBytes = limits.MaximumInputBytes,
        PackageSecurity = CreateOpenXmlPackageSecurity(limits),
        OpenSettings = CreateOpenXmlLoadSettings()
    };

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
        if (request.Registration is not null) return ConvertRegistered(request, input, diagnostics, cancellationToken);
        OfficeWorkflowRoute route = request.Route!;
        OfficeWorkflowConversionOptions settings = request.ConversionOptions ?? new();
        // The optimizer operates on unencrypted bytes; apply native output security last.
        PdfStandardEncryptionOptions? deferredEncryption = settings.CompressPdfOutput
            ? settings.GetOutputPdfOptions()?.Encryption : null;
        if (deferredEncryption != null) {
            settings = settings.Clone();
            settings.GetOutputPdfOptions()!.ClearEncryption();
        }
        PdfReadOptions? readOptions = settings.CreateReadOptions();
        cancellationToken.ThrowIfCancellationRequested();
        long maximumOutputBytes = request.Limits.MaximumOutputBytes;
        byte[] bytes;
        bool hasLoss = false;
        switch (route.Id) {
            case "book-project-epub": {
                (bytes, hasLoss) = ExportBookProject(input, maximumOutputBytes, diagnostics, cancellationToken);
                break;
            }
            case "doc-pdf": {
                using var source = new MemoryStream(input, writable: false);
                PdfDocumentConversionResult conversion = LegacyDocPdfConverter.ToPdfDocumentResult(source,
                    pdfOptions: settings.Word,
                    importOptions: new OfficeIMO.Word.LegacyDoc.LegacyDocImportOptions {
                        MaxInputBytes = (int)Math.Min(int.MaxValue, request.Limits.MaximumInputBytes)
                    }, lossPolicy: settings.LegacyDocLossPolicy, cancellationToken: cancellationToken);
                bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss;
                AddPdfWarnings(conversion.Warnings, diagnostics);
                foreach (IOfficeConversionReport report in conversion.SourceConversionReports)
                    foreach (OfficeConversionFidelityDiagnostic finding in report.FidelityDiagnostics)
                        diagnostics.Add(new OfficeWorkflowDiagnostic(finding.Code, finding.Message,
                            finding.LossKind == OfficeConversionLossKind.None ? OfficeWorkflowDiagnosticSeverity.Information : OfficeWorkflowDiagnosticSeverity.Warning,
                            "import", new Dictionary<string, string> { ["source"] = finding.Source, ["lossKind"] = finding.LossKind.ToString() }));
                break;
            }
            case "txt-pdf": {
                PdfDocumentConversionResult conversion = PdfPlainTextConverter.ToPdfDocumentResult(input, settings.PlainText, cancellationToken);
                bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss;
                AddPdfWarnings(conversion.Warnings, diagnostics);
                break;
            }
            case "docx-pdf":
                using (var source = new MemoryStream(input, writable: false))
                using (WordDocument document = settings.SourcePassword == null
                    ? WordDocument.LoadAsync(source, CreateWordLoadOptions(request.Limits), cancellationToken).GetAwaiter().GetResult()
                    : WordDocument.LoadEncrypted(source, settings.SourcePassword, CreateWordLoadOptions(request.Limits))) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var options = settings.Word ?? new WordToPdfOptions().UseProfile(ToPdfExportProfile(request.OutputProfile));
                    PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options, cancellationToken);
                    bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                    hasLoss = conversion.HasLoss;
                    AddPdfWarnings(conversion.Warnings, diagnostics);
                }
                break;
            case "xlsx-pdf":
                using (var source = new MemoryStream(input, writable: false))
                using (ExcelDocument document = (settings.SourcePassword == null
                    ? ExcelDocument.LoadAsync(source, CreateExcelLoadOptions(request.Limits), cancellationToken)
                    : ExcelDocument.LoadEncryptedAsync(source, settings.SourcePassword, CreateExcelLoadOptions(request.Limits), cancellationToken)).GetAwaiter().GetResult()) {
                    var options = settings.Excel ?? new ExcelToPdfOptions().UseProfile(ToPdfExportProfile(request.OutputProfile));
                    if (settings.WorksheetLayout.HasValue) options.WorksheetLayout = settings.WorksheetLayout.Value;
                    PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options, cancellationToken);
                    bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                    hasLoss = conversion.HasLoss;
                    AddPdfWarnings(conversion.Warnings, diagnostics);
                }
                break;
            case "pptx-pdf":
                using (var source = new MemoryStream(input, writable: false))
                using (PowerPointPresentation document = (settings.SourcePassword == null
                    ? PowerPointPresentation.LoadAsync(source, CreatePowerPointLoadOptions(request.Limits), cancellationToken)
                    : PowerPointPresentation.LoadEncryptedAsync(source, settings.SourcePassword, CreatePowerPointLoadOptions(request.Limits), cancellationToken)).GetAwaiter().GetResult()) {
                    var options = settings.PowerPoint ?? new PowerPointToPdfOptions().UseProfile(ToPdfExportProfile(request.OutputProfile));
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
                    htmlResourceSnapshots ?? (request.InputStream is null ? null : new Dictionary<string, byte[]>()), settings.Html);
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
            case "markdown-pdf": {
                var options = settings.Markdown ?? new MarkdownToPdfOptions();
                options.BaseDirectory ??= Path.GetDirectoryName(Path.GetFullPath(request.InputPath));
                using var source = new MemoryStream(input, writable: false);
                PdfDocumentConversionResult conversion = MarkdownDoc.Load(source)
                    .ToPdfDocumentResult(options, cancellationToken);
                bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss;
                AddPdfWarnings(conversion.Warnings, diagnostics);
                break;
            }
            case "rtf-pdf": {
                RtfReadResult imported = RtfDocument.LoadResult(input, cancellationToken: cancellationToken);
                foreach (var finding in imported.Diagnostics) diagnostics.Add(new OfficeWorkflowDiagnostic(finding.Code, finding.Message,
                    finding.Severity == OfficeIMO.Rtf.Diagnostics.RtfDiagnosticSeverity.Info ? OfficeWorkflowDiagnosticSeverity.Information :
                    finding.Severity == OfficeIMO.Rtf.Diagnostics.RtfDiagnosticSeverity.Error ? OfficeWorkflowDiagnosticSeverity.Error : OfficeWorkflowDiagnosticSeverity.Warning,
                    "import", new Dictionary<string, string> { ["source"] = "RTF", ["position"] = finding.Position.ToString(System.Globalization.CultureInfo.InvariantCulture) }));
                if (imported.Diagnostics.Any(finding => finding.Severity == OfficeIMO.Rtf.Diagnostics.RtfDiagnosticSeverity.Error))
                    throw new InvalidDataException("RTF import reported invalid or unsupported content that prevents conversion.");
                PdfDocumentConversionResult conversion = imported.Document.ToPdfDocumentResult(settings.Rtf, cancellationToken);
                bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss || imported.Diagnostics.Any(finding => finding.Severity != OfficeIMO.Rtf.Diagnostics.RtfDiagnosticSeverity.Info);
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
            compression.KeepOriginalWhenNotSmaller = true;
            compression.CancellationToken = cancellationToken;
            compression.MaximumOutputBytes = maximumOutputBytes;
            PdfOptimizationActionResult optimized = PdfDocument.Load(bytes, request.OutputPdfLoadOptions).Optimization.Apply(compression);
            if (!optimized.PreservationReport.IsPreserved) throw new InvalidOperationException("PDF compression did not preserve the converted document.");
            bytes = optimized.Bytes;
            diagnostics.Add(new OfficeWorkflowDiagnostic("PdfOutputCompression", "Verified lossless PDF compression completed; saved " + optimized.SavedBytes + " bytes.",
                OfficeWorkflowDiagnosticSeverity.Information, "convert"));
        }
        if (deferredEncryption != null) {
            PdfSecurityMutationResult encrypted = PdfSecurityEditor.Encrypt(bytes, deferredEncryption,
                maximumOutputBytes: maximumOutputBytes, cancellationToken: cancellationToken);
            if (!encrypted.PreservationReport.IsPreserved)
                throw new InvalidOperationException("PDF encryption did not preserve the converted document.");
            bytes = encrypted.Pdf;
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
