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
        OfficeWorkflowRoute route = request.Route!;
        cancellationToken.ThrowIfCancellationRequested();
        long maximumOutputBytes = request.Limits.MaximumOutputBytes;
        byte[] bytes;
        bool hasLoss = false;
        switch (route.Id) {
            case "docx-pdf":
                using (WordDocument document = WordDocument.Load(request.InputPath)) {
                    var options = new WordPdfSaveOptions { CancellationToken = cancellationToken };
                    options.UseProfile(ToPdfExportProfile(request.OutputProfile));
                    PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options);
                    bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                    hasLoss = conversion.HasLoss;
                    AddPdfWarnings(conversion.Warnings, diagnostics);
                }
                break;
            case "xlsx-pdf":
                using (ExcelDocument document = ExcelDocument.Load(request.InputPath)) {
                    var options = new ExcelPdfSaveOptions { CancellationToken = cancellationToken };
                    options.UseProfile(ToPdfExportProfile(request.OutputProfile));
                    PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options);
                    bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                    hasLoss = conversion.HasLoss;
                    AddPdfWarnings(conversion.Warnings, diagnostics);
                }
                break;
            case "pptx-pdf":
                using (PowerPointPresentation document = PowerPointPresentation.Load(request.InputPath)) {
                    var options = new PowerPointPdfSaveOptions { CancellationToken = cancellationToken };
                    options.UseProfile(ToPdfExportProfile(request.OutputProfile));
                    PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options);
                    bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                    hasLoss = conversion.HasLoss;
                    AddPdfWarnings(conversion.Warnings, diagnostics);
                }
                break;
            case "html-pdf": {
                string html = File.ReadAllText(request.InputPath, Encoding.UTF8);
                PdfDocumentConversionResult conversion = HtmlConversionDocument.Parse(html)
                    .ToPdfDocumentResultAsync(new HtmlPdfSaveOptions(), cancellationToken)
                    .GetAwaiter()
                    .GetResult();
                bytes = SerializePdfConversion(conversion, maximumOutputBytes, cancellationToken);
                hasLoss = conversion.HasLoss;
                AddPdfWarnings(conversion.Warnings, diagnostics);
                break;
            }
            case "pdf-docx": {
                PdfDocument pdf = PdfDocument.Load(request.InputPath, request.PdfLoadOptions);
                PdfWordConversionResult conversion = pdf.ToWordDocumentResult(new PdfWordImportOptions {
                    CancellationToken = cancellationToken
                });
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
                PdfDocument pdf = PdfDocument.Load(request.InputPath, request.PdfLoadOptions);
                PdfExcelTableImportResult conversion = pdf.ImportTablesToExcelDocumentResult(new PdfExcelTableImportOptions {
                    CancellationToken = cancellationToken
                });
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
                PdfDocument pdf = PdfDocument.Load(request.InputPath, request.PdfLoadOptions);
                PdfPowerPointConversionResult conversion = pdf.ToPowerPointPresentationResult(
                    CreatePowerPointImportOptions(cancellationToken));
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
                PdfDocument pdf = PdfDocument.Load(request.InputPath, request.PdfLoadOptions);
                PdfHtmlConversionResult conversion = pdf.ToHtmlResult(new PdfHtmlSaveOptions {
                    Profile = PdfHtmlProfile.PositionedReview,
                    IncludeLinkAnnotations = true,
                    IncludeFormWidgets = true,
                    CancellationToken = cancellationToken
                });
                bytes = EncodeUtf8Bounded(conversion.Value, maximumOutputBytes);
                hasLoss = conversion.HasLoss;
                AddMessages(conversion.Report.Warnings.Select(static warning => warning.ToString()), hasLoss, diagnostics);
                break;
            }
            default:
                throw new NotSupportedException("The conversion route '" + route.Id + "' is not implemented by the local runner.");
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
}
