using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[] Bytes, OfficeWorkflowConversionEvidence Evidence) ConvertTier1(ValidatedRequest request,
        byte[] input, CancellationToken token, bool emitHtmlTaggedStructure,
        IReadOnlyDictionary<string, byte[]>? htmlResourceSnapshots) {
        OfficeConversionOptions options = request.UnifiedConversionOptions?.Clone() ?? new();
        options.RouteId = request.Route!.Id;
        options.SourceName = Path.GetFileName(request.InputPath);
        options.TargetName = "output" + request.Route.TargetExtension;
        options.MaximumInputBytes = request.Limits.MaximumInputBytes;
        options.MaximumOutputBytes = request.Limits.MaximumOutputBytes;
        options.MaximumXmlCharactersInPart = request.Limits.MaximumXmlCharactersInPart;
        options.OutputProfile = request.OutputProfile;
        options.PdfLoad = (options.PdfLoad ?? request.PdfLoadOptions).Snapshot(request.Limits.MaximumInputBytes,
            options.PdfLoad?.Password ?? request.PdfLoadOptions.Password);
        if (!emitHtmlTaggedStructure) {
            options.Html ??= new HtmlToPdfOptions();
            options.Html.PdfOptions.SetTaggedStructureMode(PdfTaggedStructureMode.None);
        }
        options.ResourceDirectory = Path.GetDirectoryName(Path.GetFullPath(request.InputPath));
        options.ResourceSnapshots = htmlResourceSnapshots ?? (request.InputStream == null ? null : new Dictionary<string, byte[]>());
        using var source = new MemoryStream(input, false);
        using var output = new OfficeWorkflowBoundedMemoryStream(request.Limits.MaximumOutputBytes);
        try {
            OfficeConversionReport report = OfficeConverter.ConvertAsync(source, OfficeConverter.FormatForExtension(Path.GetExtension(options.SourceName)),
                output, OfficeConverter.FormatForExtension(request.Route.TargetExtension), options, token).GetAwaiter().GetResult();
            return (output.ToArray(), new OfficeWorkflowConversionEvidence(report));
        } catch (OfficeConversionException error) {
            throw new WorkflowConversionFailureException(error, new OfficeWorkflowConversionEvidence(error.Report), importDiagnosticSources: new HashSet<string>(new[] { "Word", "Excel", "PowerPoint" }));
        }
    }
}
