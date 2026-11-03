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
                using var output = new MemoryStream();
                report = source.SaveAsEpub(output, new ChmEpubExportOptions { ConversionOptions = conversion,
                    SaveOptions = new OfficeIMO.Epub.EpubSaveOptions { MaxOutputBytes = conversion.MaxOutputBytes } }, token);
                bytes = output.ToArray(); break;
            }
            default: {
                using var output = new MemoryStream();
                report = source.SaveAsMarkdown(output, new ChmMarkdownExportOptions { ConversionOptions = conversion }, token);
                bytes = output.ToArray(); break;
            }
        }
        var evidence = new OfficeWorkflowConversionEvidence(report, new Dictionary<string, string> {
            ["sourceContainer"] = "ITSF", ["topicCount"] = source.Topics.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["resourcePolicy"] = "Embedded archive resources only"
        });
        return (bytes, evidence);
    }
}
