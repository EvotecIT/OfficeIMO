using OfficeIMO.Chm;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[] Bytes, OfficeWorkflowConversionEvidence Evidence) ConvertChm(ValidatedRequest request, byte[] input,
        OfficeWorkflowRenderingOptions settings, CancellationToken token, bool emitHtmlTaggedStructure) {
        var reading = new ChmLoadOptions(); reading.MaxInputBytes = Math.Min(reading.MaxInputBytes, request.Limits.MaximumInputBytes);
        ChmDocument source = ChmDocument.Load(input, reading, token);
        ChmConversionOptions conversion = settings.Chm?.Clone() ?? new ChmConversionOptions();
        conversion.MaxOutputBytes = Math.Min(conversion.MaxOutputBytes, request.Limits.MaximumOutputBytes);
        byte[] bytes; OfficeConversionReport report;
        switch (request.Route!.Id) {
            case "chm-pdf": {
                var rendering = settings.Html?.ClonePdf() ?? new OfficeIMO.Html.Pdf.HtmlPdfExportOptions();
                if (!emitHtmlTaggedStructure) rendering.PdfOptions.SetTaggedStructureMode(OfficeIMO.Pdf.PdfTaggedStructureMode.None);
                using (OfficeIMO.Pdf.PdfDocument pdf = source.ToPdf(new ChmPdfExportOptions { ConversionOptions = conversion, PdfOptions = rendering }, out report, token))
                    bytes = pdf.ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }, token);
                if (bytes.LongLength > conversion.MaxOutputBytes) throw new InvalidDataException("CHM PDF output exceeds MaxOutputBytes.");
                break;
            }
            case "chm-epub": {
                var result = source.ToEpubBytesResult(conversion, writeOptions: new OfficeIMO.Epub.EpubSaveOptions { MaxOutputBytes = conversion.MaxOutputBytes }, cancellationToken: token);
                bytes = result.Value; report = result.Report; break;
            }
            default: {
                var result = source.ToMarkdownResult(conversion, cancellationToken: token); bytes = EncodeUtf8Bounded(result.Value, conversion.MaxOutputBytes); report = result.Report; break;
            }
        }
        var evidence = new OfficeWorkflowConversionEvidence(report, new Dictionary<string, string> {
            ["sourceContainer"] = "ITSF", ["topicCount"] = source.Topics.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["resourcePolicy"] = "Embedded archive resources only"
        });
        return (bytes, evidence);
    }
}
