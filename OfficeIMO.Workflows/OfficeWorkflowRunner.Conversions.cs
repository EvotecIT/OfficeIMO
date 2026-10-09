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
using OfficeIMO.Xps;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static OfficePackageSecurityOptions CreateOpenXmlPackageSecurity(OfficeWorkflowLimits limits) {
        OfficePackageSecurityOptions security = OfficePackageSecurityOptions.SecureDefaults;
        security.MaxPackageBytes = limits.MaximumInputBytes;
        security.MaxXmlCharactersInPart = limits.MaximumXmlCharactersInPart;
        return security;
    }

    private static OfficeOpenXmlLoadSettings CreateOpenXmlLoadSettings(OfficeWorkflowLimits limits) => new() {
        MaxCharactersInPart = limits.MaximumXmlCharactersInPart
    };

    private static WordLoadOptions CreateWordLoadOptions(OfficeWorkflowLimits limits) => new() {
        AccessMode = DocumentAccessMode.ReadOnly,
        MaxInputBytes = limits.MaximumInputBytes,
        PackageSecurity = CreateOpenXmlPackageSecurity(limits),
        OpenSettings = CreateOpenXmlLoadSettings(limits)
    };

    private static ExcelLoadOptions CreateExcelLoadOptions(OfficeWorkflowLimits limits) => new() {
        AccessMode = DocumentAccessMode.ReadOnly,
        MaxInputBytes = limits.MaximumInputBytes,
        PackageSecurity = CreateOpenXmlPackageSecurity(limits),
        OpenSettings = CreateOpenXmlLoadSettings(limits)
    };

    private static PowerPointLoadOptions CreatePowerPointLoadOptions(OfficeWorkflowLimits limits) => new() {
        AccessMode = DocumentAccessMode.ReadOnly,
        MaxInputBytes = limits.MaximumInputBytes,
        PackageSecurity = CreateOpenXmlPackageSecurity(limits),
        OpenSettings = CreateOpenXmlLoadSettings(limits)
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
        OfficeWorkflowConversionEvidence? evidence = null;
        HashSet<string>? importDiagnosticSources = null;
        switch (route.Id) {
            case "publisher-pdf": {
                (bytes, evidence) = ConvertPublisher(request, input, settings, cancellationToken);
                hasLoss = evidence.HasLoss;
                break;
            }
            case "visio-pdf": {
                (bytes, evidence) = ConvertVisio(request, input, settings, cancellationToken);
                hasLoss = evidence.HasLoss;
                break;
            }
            case "odg-pdf": {
                (bytes, evidence) = ConvertDraw(request, input, settings, cancellationToken);
                hasLoss = evidence.HasLoss;
                break;
            }
            case "xps-pdf": {
                var limits = new XpsReadOptions();
                limits.MaximumInputBytes = (int)Math.Min(limits.MaximumInputBytes, request.Limits.MaximumInputBytes);
                var document = XpsDocument.Load(input, limits, cancellationToken);
                using var output = new OfficeWorkflowBoundedMemoryStream(maximumOutputBytes);
                document.SavePdf(output, settings.Xps, cancellationToken);
                bytes = output.ToArray();
                break;
            }
            case "book-project-epub": {
                (bytes, hasLoss) = ExportBookProject(input, maximumOutputBytes, diagnostics, cancellationToken);
                break;
            }
            case "doc-pdf": {
                using var source = new MemoryStream(input, writable: false);
                PdfDocumentConversionResult conversion;
                try {
                    conversion = LegacyDocPdfConverter.ToPdfDocumentResult(source,
                        pdfOptions: settings.Word,
                        importOptions: new OfficeIMO.Word.LegacyDoc.LegacyDocImportOptions {
                            MaxInputBytes = (int)Math.Min(int.MaxValue, request.Limits.MaximumInputBytes)
                        }, lossPolicy: settings.LegacyDocLossPolicy, cancellationToken: cancellationToken);
                } catch (OfficeConversionException exception) {
                    var rejected = new OfficeWorkflowConversionEvidence(exception.Report);
                    var sources = new HashSet<string>(rejected.FidelityDiagnostics.Select(finding => finding.Source), StringComparer.Ordinal);
                    throw new WorkflowConversionFailureException(exception.InnerException ?? exception, rejected, importDiagnosticSources: sources);
                }
                importDiagnosticSources = new HashSet<string>(conversion.SourceConversionReports
                    .SelectMany(report => report.FidelityDiagnostics).Select(finding => finding.Source), StringComparer.Ordinal);
                (bytes, evidence) = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken, importDiagnosticSources: importDiagnosticSources);
                hasLoss = conversion.HasLoss;
                break;
            }
            case "txt-pdf": {
                PdfDocumentConversionResult conversion = PdfPlainTextConverter.ToPdfDocumentResult(input, settings.PlainText, cancellationToken);
                (bytes, evidence) = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss;
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
                    (bytes, evidence) = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                    hasLoss = conversion.HasLoss;
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
                    (bytes, evidence) = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                    hasLoss = conversion.HasLoss;
                }
                break;
            case "pptx-pdf":
                using (var source = new MemoryStream(input, writable: false))
                using (PowerPointPresentation document = (settings.SourcePassword == null
                    ? PowerPointPresentation.LoadAsync(source, CreatePowerPointLoadOptions(request.Limits), cancellationToken)
                    : PowerPointPresentation.LoadEncryptedAsync(source, settings.SourcePassword, CreatePowerPointLoadOptions(request.Limits), cancellationToken)).GetAwaiter().GetResult()) {
                    var options = settings.PowerPoint ?? new PowerPointToPdfOptions().UseProfile(ToPdfExportProfile(request.OutputProfile));
                    PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options, cancellationToken);
                    (bytes, evidence) = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                    hasLoss = conversion.HasLoss;
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
                (bytes, evidence) = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss;
                break;
            }
            case "markdown-pdf": {
                var options = settings.Markdown ?? new MarkdownToPdfOptions();
                options.BaseDirectory ??= Path.GetDirectoryName(Path.GetFullPath(request.InputPath));
                using var source = new MemoryStream(input, writable: false);
                PdfDocumentConversionResult conversion = MarkdownDoc.Load(source)
                    .ToPdfDocumentResult(options, cancellationToken);
                (bytes, evidence) = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss;
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
                (bytes, evidence) = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss || imported.Diagnostics.Any(finding => finding.Severity != OfficeIMO.Rtf.Diagnostics.RtfDiagnosticSeverity.Info);
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
                (bytes, evidence) = SerializeEditableConversion(conversion.Report,
                    stream => document.SaveAsync(stream, cancellationToken).GetAwaiter().GetResult(), maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss;
                break;
            }
            case "pdf-xlsx": {
                PdfDocument pdf = PdfDocument.Load(input, request.PdfLoadOptions);
                PdfExcelTableImportResult conversion = pdf.ImportTablesToExcelDocumentResult(new PdfTablesToExcelOptions { ReadOptions = readOptions }, cancellationToken);
                using ExcelDocument document = conversion.Value;
                (bytes, evidence) = SerializeEditableConversion(conversion.Report,
                    stream => document.SaveAsync(stream, cancellationToken).GetAwaiter().GetResult(), maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss || conversion.HasOmittedPageContent;
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
                (bytes, evidence) = SerializeEditableConversion(conversion.Report,
                    stream => document.SaveAsync(stream, cancellationToken).GetAwaiter().GetResult(), maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss || conversion.HasOmittedPageContent;
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
                (bytes, evidence) = SerializeReportedConversion(conversion.Report,
                    () => EncodeUtf8Bounded(conversion.Value, maximumOutputBytes), cancellationToken);
                hasLoss = conversion.HasLoss;
                break;
            }
            default:
                throw new NotSupportedException("The conversion route '" + route.Id + "' is not implemented by the local runner.");
        }

        try {
            if (evidence != null) {
                AddConversionDiagnostics(evidence, diagnostics, importDiagnosticSources);
                hasLoss |= evidence.HasLoss;
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
            return new OperationArtifact(bytes, summary, null, ConversionEvidence: evidence);
        } catch (OperationCanceledException exception) when (cancellationToken.IsCancellationRequested && evidence != null
            && exception is not WorkflowConversionCancellationException) {
            throw new WorkflowConversionCancellationException(exception, evidence, diagnosticsAdded: true, importDiagnosticSources: importDiagnosticSources);
        } catch (Exception exception) when (evidence != null && exception is not WorkflowConversionFailureException and not OperationCanceledException
            and not OutOfMemoryException and not StackOverflowException) {
            throw new WorkflowConversionFailureException(exception, evidence, diagnosticsAdded: true, importDiagnosticSources: importDiagnosticSources);
        }

    }

    private static void AddConversionDiagnostics(OfficeWorkflowConversionEvidence evidence, List<OfficeWorkflowDiagnostic> diagnostics,
        ISet<string>? importDiagnosticSources = null) {
        foreach (OfficeConversionFidelityDiagnostic finding in evidence.FidelityDiagnostics) {
            var details = new Dictionary<string, string> { ["source"] = finding.Source, ["lossKind"] = finding.LossKind.ToString() };
            if (finding.Location != null) details["location"] = finding.Location;
            diagnostics.Add(new OfficeWorkflowDiagnostic(finding.Code, finding.Message,
                finding.LossKind == OfficeConversionLossKind.Failure ? OfficeWorkflowDiagnosticSeverity.Error
                    : finding.LossKind == OfficeConversionLossKind.None ? OfficeWorkflowDiagnosticSeverity.Information : OfficeWorkflowDiagnosticSeverity.Warning,
                importDiagnosticSources?.Contains(finding.Source) == true ? "import" : "convert", details));
        }
    }

    private static (byte[] Bytes, OfficeWorkflowConversionEvidence Evidence) SerializePdfConversion(PdfDocumentConversionResult conversion, long maximumOutputBytes,
        CancellationToken cancellationToken, IReadOnlyDictionary<string, string>? facts = null, ISet<string>? importDiagnosticSources = null) {
        using var stream = new OfficeWorkflowBoundedMemoryStream(maximumOutputBytes);
        PdfSaveResult? saved = null;
        try {
            saved = conversion.SaveResultAsync(stream, cancellationToken).GetAwaiter().GetResult();
            if (saved.Exception is OutOfMemoryException or StackOverflowException)
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(saved.Exception).Throw();
            cancellationToken.ThrowIfCancellationRequested();
            if (!saved.Succeeded) {
                throw new WorkflowConversionFailureException(saved.Exception!,
                    new OfficeWorkflowConversionEvidence(saved.ConversionReports, facts ?? new Dictionary<string, string>()), importDiagnosticSources: importDiagnosticSources);
            }
            return (stream.ToArray(), new OfficeWorkflowConversionEvidence(saved.ConversionReports,
                facts ?? new Dictionary<string, string>()));
        } catch (OperationCanceledException exception) when (cancellationToken.IsCancellationRequested) {
            throw new WorkflowConversionCancellationException(exception,
                new OfficeWorkflowConversionEvidence(saved?.ConversionReports ?? conversion.ConversionReports,
                    facts ?? new Dictionary<string, string>()), importDiagnosticSources: importDiagnosticSources);
        }
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
