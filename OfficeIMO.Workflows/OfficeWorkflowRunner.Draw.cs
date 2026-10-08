using System.Globalization;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odg.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[], OfficeWorkflowConversionEvidence) ConvertDraw(ValidatedRequest request, byte[] input,
        OfficeWorkflowConversionOptions settings, List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        using var stream = new MemoryStream(input, writable: false);
        var loadOptions = new OdfLoadOptions {
            MaxPackageBytes = request.Limits.MaximumInputBytes,
            MaxXmlCharacters = request.Limits.MaximumXmlCharactersInPart
        };
        bool flat = Path.GetExtension(request.InputPath).Equals(".fodg", StringComparison.OrdinalIgnoreCase);
        OdgDocument source = flat ? OdgDocument.LoadFlatXml(stream, loadOptions) : OdgDocument.Load(stream, loadOptions);
        var facts = new Dictionary<string, string> {
            ["sourceFormat"] = flat ? "FODG" : "ODG",
            ["sourcePages"] = source.Pages.Count.ToString(CultureInfo.InvariantCulture),
            ["layerVisibility"] = settings.Draw?.ForPrint == false ? "screen" : "print",
            ["pagePolicy"] = "source-dimensions"
        };
        PdfDocumentConversionResult? conversion = null;
        try {
            conversion = source.ToPdfDocumentResult(settings.Draw, token);
            byte[] bytes = SerializePdfConversion(conversion, request.Limits.MaximumOutputBytes, token, facts);
            var evidence = new OfficeWorkflowConversionEvidence(conversion.ConversionReports, facts);
            if (settings.RequireNoLoss) evidence.RequireNoLoss();
            AddConversionDiagnostics(evidence, diagnostics);
            return (bytes, evidence);
        } catch (OdfConversionLossException exception) {
            throw new WorkflowConversionFailureException(exception, new OfficeWorkflowConversionEvidence(exception.Report, facts));
        } catch (Exception exception) when (conversion != null && exception is not WorkflowConversionFailureException and not OperationCanceledException
            and not OutOfMemoryException and not StackOverflowException) {
            throw new WorkflowConversionFailureException(exception, new OfficeWorkflowConversionEvidence(conversion.ConversionReports, facts));
        }
    }

    private static void AddConversionDiagnostics(OfficeWorkflowConversionEvidence evidence, List<OfficeWorkflowDiagnostic> diagnostics) {
        foreach (OfficeConversionFidelityDiagnostic finding in evidence.FidelityDiagnostics) {
            var details = new Dictionary<string, string> { ["source"] = finding.Source, ["lossKind"] = finding.LossKind.ToString() };
            if (finding.Location != null) details["location"] = finding.Location;
            diagnostics.Add(new OfficeWorkflowDiagnostic(finding.Code, finding.Message,
                finding.LossKind == OfficeConversionLossKind.Failure ? OfficeWorkflowDiagnosticSeverity.Error
                    : finding.LossKind == OfficeConversionLossKind.None ? OfficeWorkflowDiagnosticSeverity.Information : OfficeWorkflowDiagnosticSeverity.Warning,
                "convert", details));
        }
    }
}
