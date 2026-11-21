using OfficeIMO.Chm;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[] Bytes, OfficeWorkflowConversionEvidence Evidence) ConvertChm(ValidatedRequest request, byte[] input,
        OfficeWorkflowRenderingOptions settings, CancellationToken token, bool emitHtmlTaggedStructure) {
        var reading = new ChmLoadOptions(); reading.MaxInputBytes = Math.Min(reading.MaxInputBytes, request.Limits.MaximumInputBytes);
        ChmDocument source = ChmDocument.Load(input, reading, token);
        ChmConversionOptions conversion = settings.Chm?.Clone() ?? new ChmConversionOptions();
        conversion.MaxOutputBytes = Math.Min(conversion.MaxOutputBytes, request.Limits.MaximumOutputBytes);
        var facts = new Dictionary<string, string> {
            ["sourceContainer"] = "ITSF", ["topicCount"] = source.Topics.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["resourcePolicy"] = "Embedded archive resources only"
        };
        try {
            switch (request.Route!.Id) {
                case "chm-pdf": {
                    var rendering = settings.Html?.ClonePdf() ?? new OfficeIMO.Html.Pdf.HtmlPdfExportOptions();
                    if (!emitHtmlTaggedStructure) rendering.PdfOptions.SetTaggedStructureMode(OfficeIMO.Pdf.PdfTaggedStructureMode.None);
                    using OfficeIMO.Pdf.PdfDocument pdf = source.ToPdf(new ChmPdfExportOptions {
                        ConversionOptions = conversion, PdfOptions = rendering
                    }, out OfficeConversionReport report, token);
                    return SerializeChmConversion(report, () => {
                        byte[] bytes = pdf.ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }, token);
                        if (bytes.LongLength > conversion.MaxOutputBytes)
                            throw new InvalidDataException("CHM PDF output exceeds MaxOutputBytes.");
                        return bytes;
                    }, token, facts);
                }
                case "chm-epub": {
                    using var output = new MemoryStream();
                    OfficeConversionReport report = source.SaveAsEpub(output, new ChmEpubExportOptions {
                        ConversionOptions = conversion,
                        SaveOptions = new OfficeIMO.Epub.EpubSaveOptions { MaxOutputBytes = conversion.MaxOutputBytes }
                    }, token);
                    return SerializeChmConversion(report, output.ToArray, token, facts);
                }
                default: {
                    using var output = new MemoryStream();
                    OfficeConversionReport report = source.SaveAsMarkdown(output,
                        new ChmMarkdownExportOptions { ConversionOptions = conversion }, token);
                    return SerializeChmConversion(report, output.ToArray, token, facts);
                }
            }
        } catch (OfficeConversionException exception) {
            throw new WorkflowConversionFailureException(exception, new OfficeWorkflowConversionEvidence(exception.Report, facts));
        }
    }

    private static (byte[] Bytes, OfficeWorkflowConversionEvidence Evidence) SerializeChmConversion(
        OfficeConversionReport report, Func<byte[]> serialize, CancellationToken token, IReadOnlyDictionary<string, string> facts) =>
        SerializeReportedConversion(report, () => {
            if (report.Diagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Failure))
                throw new OfficeConversionException("The CHM conversion failed. Inspect its report before publishing.", report);
            return serialize();
        }, token, facts);
}
