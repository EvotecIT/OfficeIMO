using OfficeIMO.Chm;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[] Bytes, OfficeWorkflowConversionEvidence Evidence) ConvertChm(ValidatedRequest request, byte[] input,
        OfficeWorkflowConversionOptions settings, List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token, bool emitHtmlTaggedStructure) {
        var reading = new ChmReadOptions(); reading.MaxInputBytes = Math.Min(reading.MaxInputBytes, request.Limits.MaximumInputBytes);
        ChmDocument source = ChmDocument.Load(input, reading, token);
        ChmConversionOptions conversion = settings.Chm?.Clone() ?? new ChmConversionOptions();
        conversion.MaxOutputBytes = Math.Min(conversion.MaxOutputBytes, request.Limits.MaximumOutputBytes);
        byte[] bytes; IOfficeConversionReport report;
        switch (request.Route!.Id) {
            case "chm-pdf": {
                var rendering = settings.Html?.ClonePdf() ?? new OfficeIMO.Html.Pdf.HtmlToPdfOptions();
                if (!emitHtmlTaggedStructure) rendering.PdfOptions.SetTaggedStructureMode(OfficeIMO.Pdf.PdfTaggedStructureMode.None);
                var result = source.ToPdfBytesResult(conversion, rendering, token); bytes = result.RequireValue(); report = result.Report; break;
            }
            case "chm-epub": {
                var result = source.ToEpubBytesResult(conversion, writeOptions: new OfficeIMO.Epub.EpubWriteOptions { MaxOutputBytes = conversion.MaxOutputBytes }, cancellationToken: token);
                bytes = result.RequireValue(); report = result.Report; break;
            }
            default: {
                var result = source.ToMarkdownResult(conversion, cancellationToken: token); bytes = EncodeUtf8Bounded(result.RequireValue(), conversion.MaxOutputBytes); report = result.Report; break;
            }
        }
        var evidence = new OfficeWorkflowConversionEvidence(report, new Dictionary<string, string> {
            ["sourceContainer"] = "ITSF", ["topicCount"] = source.Topics.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["resourcePolicy"] = "Embedded archive resources only"
        });
        foreach (OfficeConversionFidelityDiagnostic finding in evidence.FidelityDiagnostics)
            diagnostics.Add(new OfficeWorkflowDiagnostic(finding.Code, finding.Message,
                finding.LossKind == OfficeConversionLossKind.None ? OfficeWorkflowDiagnosticSeverity.Information : OfficeWorkflowDiagnosticSeverity.Warning,
                "convert", new Dictionary<string, string> { ["source"] = finding.Source, ["lossKind"] = finding.LossKind.ToString() }));
        return (bytes, evidence);
    }
}
