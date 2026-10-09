using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[], OfficeWorkflowConversionEvidence) ConvertPublisher(ValidatedRequest request, byte[] input,
        OfficeWorkflowConversionOptions settings, List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        PublisherReadOptions read = settings.PublisherRead?.Clone() ?? new PublisherReadOptions();
        read.Limits.MaxInputBytes = (int)Math.Min(read.Limits.MaxInputBytes, request.Limits.MaximumInputBytes);
        PublisherDocument source = PublisherDocument.Load(input, read, token);
        var facts = new Dictionary<string, string> {
            ["sourceFormat"] = "PUB",
            ["sourcePages"] = source.Pages.Count.ToString(CultureInfo.InvariantCulture),
            ["sourceStories"] = source.TextStories.Count.ToString(CultureInfo.InvariantCulture),
            ["pagePolicy"] = "source-dimensions"
        };
        PdfDocumentConversionResult? conversion = null;
        try {
            conversion = source.ToPdfDocumentResult(settings.PublisherPdf, token);
            var (bytes, evidence) = SerializePdfConversion(conversion, request.Limits.MaximumOutputBytes, token, facts);
            if (settings.RequireNoLoss) evidence.RequireNoLoss();
            AddConversionDiagnostics(evidence, diagnostics);
            return (bytes, evidence);
        } catch (OperationCanceledException exception) when (token.IsCancellationRequested && exception is not WorkflowConversionCancellationException) {
            IReadOnlyList<IOfficeConversionReport> reports = conversion?.ConversionReports ?? new IOfficeConversionReport[] { source.ReadReport };
            throw new WorkflowConversionCancellationException(exception, new OfficeWorkflowConversionEvidence(reports, facts));
        } catch (OfficeConversionException exception) {
            var reports = new List<IOfficeConversionReport> { source.ReadReport };
            reports.Add(exception.Report);
            throw new WorkflowConversionFailureException(exception, new OfficeWorkflowConversionEvidence(reports, facts));
        } catch (Exception exception) when (conversion != null && exception is not WorkflowConversionFailureException and not OperationCanceledException
            and not OutOfMemoryException and not StackOverflowException) {
            throw new WorkflowConversionFailureException(exception, new OfficeWorkflowConversionEvidence(conversion.ConversionReports, facts));
        }
    }
}
