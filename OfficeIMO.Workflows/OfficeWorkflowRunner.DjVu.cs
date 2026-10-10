using System.Globalization;
using OfficeIMO.DjVu;
using OfficeIMO.DjVu.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[], OfficeWorkflowConversionEvidence) ConvertDjVu(ValidatedRequest request, byte[] input,
        OfficeWorkflowConversionOptions settings, List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        var document = DjVuDocument.Load(input, new DjVuReadOptions {
            MaxSourceBytes = Math.Min(256L * 1024 * 1024, request.Limits.MaximumInputBytes)
        }, token);
        var options = settings.DjVu?.Clone() ?? new DjVuToPdfOptions();
        options.MaxPdfBytes = Math.Min(options.MaxPdfBytes, request.Limits.MaximumOutputBytes);
        var facts = new Dictionary<string, string> {
            ["sourceFormat"] = "DjVu", ["sourcePages"] = document.Pages.Count.ToString(CultureInfo.InvariantCulture),
            ["sourceSha256"] = document.SourceSha256, ["pagePolicy"] = "source-dimensions", ["textPolicy"] = "stored-text"
        };
        PdfDocumentConversionResult? conversion = null;
        try {
            conversion = document.ToPdfDocumentResult(options, token);
            byte[] bytes = SerializePdfConversion(conversion, request.Limits.MaximumOutputBytes, token, facts);
            var evidence = new OfficeWorkflowConversionEvidence(conversion.ConversionReports, facts);
            if (settings.RequireNoLoss) evidence.RequireNoLoss();
            AddConversionDiagnostics(evidence, diagnostics);
            return (bytes, evidence);
        } catch (Exception exception) when (conversion != null && exception is not WorkflowConversionFailureException and not OperationCanceledException
            and not OutOfMemoryException and not StackOverflowException) {
            throw new WorkflowConversionFailureException(exception, new OfficeWorkflowConversionEvidence(conversion.ConversionReports, facts));
        }
    }
}
