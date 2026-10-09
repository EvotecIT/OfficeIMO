using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[], OfficeWorkflowConversionEvidence) ConvertVisio(ValidatedRequest request, byte[] input,
        OfficeWorkflowConversionOptions settings, List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        using var stream = new MemoryStream(input, writable: false);
        var loadOptions = new VisioLoadOptions {
            MaxInputBytes = request.Limits.MaximumInputBytes,
            MaxLegacyXmlCharacters = request.Limits.MaximumXmlCharactersInPart,
            PackageSecurity = CreateOpenXmlPackageSecurity(request.Limits)
        };
        string extension = Path.GetExtension(request.InputPath).ToLowerInvariant();
        OfficeConversionResult<VisioDocument, VisioXmlConversionReport>? imported = extension is ".vdx" or ".vtx"
            ? VisioDocument.LoadLegacyXml(stream, extension == ".vtx" ? VisioPackageType.Template : VisioPackageType.Drawing, loadOptions, token)
            : null;
        VisioDocument source = imported?.Value ?? VisioDocument.Load(stream, loadOptions);
        var options = settings.Visio ?? new VisioToPdfOptions { Mode = VisioPdfProjectionMode.DiagramPages };
        options.SourceName ??= request.InputPath;
        var facts = new Dictionary<string, string> {
            ["sourceFormat"] = extension.TrimStart('.').ToUpperInvariant(),
            ["sourcePages"] = source.Pages.Count.ToString(CultureInfo.InvariantCulture),
            ["layerVisibility"] = (options.DrawingOptions?.LayerMode ?? VisioLayerRenderMode.Printable).ToString(),
            ["pagePolicy"] = "source-dimensions"
        };
        PdfDocumentConversionResult? conversion = null;
        try {
            conversion = source.ToPdfDocumentResult(options, token);
            if (imported != null) conversion = conversion.WithSourceConversionReport(imported.Report);
            var (bytes, evidence) = SerializePdfConversion(conversion, request.Limits.MaximumOutputBytes, token, facts);
            if (settings.RequireNoLoss) evidence.RequireNoLoss();
            AddConversionDiagnostics(evidence, diagnostics);
            return (bytes, evidence);
        } catch (OfficeConversionException exception) {
            var reports = new List<IOfficeConversionReport>();
            if (imported != null) reports.Add(imported.Report);
            reports.Add(exception.Report);
            throw new WorkflowConversionFailureException(exception, new OfficeWorkflowConversionEvidence(reports, facts));
        } catch (Exception exception) when (conversion != null && exception is not WorkflowConversionFailureException and not OperationCanceledException
            and not OutOfMemoryException and not StackOverflowException) {
            throw new WorkflowConversionFailureException(exception, new OfficeWorkflowConversionEvidence(conversion.ConversionReports, facts));
        }
    }
}
