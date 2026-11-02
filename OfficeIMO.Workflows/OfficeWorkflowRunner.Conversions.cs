using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.PowerPoint;
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
    private static OfficePackageSecurityOptions CreateOpenXmlPackageSecurity(OfficeOperationLimits limits) {
        OfficePackageSecurityOptions security = OfficePackageSecurityOptions.SecureDefaults;
        security.MaxPackageBytes = limits.MaximumInputBytes;
        security.MaxXmlCharactersInPart = limits.MaximumXmlCharactersInPart;
        return security;
    }

    private static OfficeOpenXmlLoadSettings CreateOpenXmlLoadSettings(OfficeOperationLimits limits) => new() {
        MaxCharactersInPart = limits.MaximumXmlCharactersInPart
    };

    private static WordLoadOptions CreateWordLoadOptions(OfficeOperationLimits limits) => new() {
        AccessMode = DocumentAccessMode.ReadOnly,
        MaxInputBytes = limits.MaximumInputBytes,
        PackageSecurity = CreateOpenXmlPackageSecurity(limits),
        OpenSettings = CreateOpenXmlLoadSettings(limits)
    };

    private static ExcelLoadOptions CreateExcelLoadOptions(OfficeOperationLimits limits) => new() {
        AccessMode = DocumentAccessMode.ReadOnly,
        MaxInputBytes = limits.MaximumInputBytes,
        PackageSecurity = CreateOpenXmlPackageSecurity(limits),
        OpenSettings = CreateOpenXmlLoadSettings(limits)
    };

    private static PowerPointLoadOptions CreatePowerPointLoadOptions(OfficeOperationLimits limits) => new() {
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
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (request.Registration is not null) return ConvertRegistered(request, input, diagnostics, cancellationToken);
        OfficeWorkflowRoute route = request.Route!;
        OfficeWorkflowRenderingOptions settings = request.ConversionOptions ?? new();
        bool tier1 = OfficeConverter.IsBuiltinRoute(route.Id);
        // The optimizer operates on unencrypted bytes; apply native output security last.
        PdfStandardEncryptionOptions? deferredEncryption = settings.CompressPdfOutput && !tier1
            ? settings.GetOutputPdfOptions()?.Encryption : null;
        if (deferredEncryption != null) {
            settings = settings.Clone();
            settings.GetOutputPdfOptions()!.ClearEncryption();
        }
        PdfLoadOptions? readOptions = settings.CreateReadOptions();
        cancellationToken.ThrowIfCancellationRequested();
        long maximumOutputBytes = request.Limits.MaximumOutputBytes;
        byte[] bytes;
        bool hasLoss = false;
        OfficeWorkflowConversionEvidence? evidence = null;
        HashSet<string>? importDiagnosticSources = tier1 ? new HashSet<string>(new[] { "Word", "Excel", "PowerPoint" }) : null;
        switch (route.Id) {
            case "publisher-pdf": {
                (bytes, evidence) = ConvertPublisher(request, input, settings, cancellationToken);
                hasLoss = evidence.HasLoss;
                break;
            }
            case "chm-pdf":
            case "chm-markdown":
            case "chm-epub": {
                (bytes, evidence) = ConvertChm(request, input, settings, cancellationToken, emitHtmlTaggedStructure);
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
            case "djvu-pdf": {
                (bytes, evidence) = ConvertDjVu(request, input, settings, cancellationToken);
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
            case "doc-pdf":
            case "docx-pdf":
            case "txt-pdf":
            case "xlsx-pdf":
            case "pptx-pdf":
            case "html-pdf":
            case "markdown-pdf":
            case "pdf-docx":
            case "pdf-xlsx":
            case "pdf-pptx":
            case "pdf-html": {
                (bytes, evidence) = ConvertTier1(request, input, cancellationToken, emitHtmlTaggedStructure, htmlResourceSnapshots);
                hasLoss = evidence.HasLoss;
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
                PdfDocumentConversionResult conversion = imported.Document.ConvertToPdfWithEvidence(settings.Rtf, cancellationToken);
                (bytes, evidence) = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss || imported.Diagnostics.Any(finding => finding.Severity != OfficeIMO.Rtf.Diagnostics.RtfDiagnosticSeverity.Info);
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
            if (settings.CompressPdfOutput && !tier1) {
                bytes = CompressConvertedPdfOutput(bytes, request, diagnostics, cancellationToken);
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
            saved = conversion.TrySaveWithEvidenceAsync(stream, cancellationToken).GetAwaiter().GetResult();
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

    private static byte[] CompressConvertedPdfOutput(byte[] bytes, ValidatedRequest request,
        List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken cancellationToken) {
        PdfOptimizationActionResult optimized = OfficeConverter.CompressPdfBytes(bytes, request.Limits.MaximumOutputBytes,
            request.OutputPdfLoadOptions, cancellationToken);
        diagnostics.Add(new OfficeWorkflowDiagnostic(global::OfficeIMO.OfficeDiagnosticCodes.PdfOutputCompressed, "Verified lossless PDF compression completed; saved " + optimized.SavedBytes + " bytes.",
            OfficeWorkflowDiagnosticSeverity.Information, "convert"));
        return optimized.Bytes;
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
        if (input == null) throw new ArgumentNullException(nameof(input));
        if ((inputPath == null || string.IsNullOrWhiteSpace(inputPath))) throw new ArgumentException("Input path cannot be empty.", nameof(inputPath));
        using var source = new MemoryStream(input, writable: false);
        return HtmlConversionDocument.LoadAsync(
            source,
            new HtmlConversionDocumentOptions {
                BaseUri = new Uri(Path.GetFullPath(inputPath)),
                ResourceUrlPolicy = OfficeHtmlResourceResolver.CreateResourcePolicy()
            },
            cancellationToken: cancellationToken).GetAwaiter().GetResult();
    }
}
