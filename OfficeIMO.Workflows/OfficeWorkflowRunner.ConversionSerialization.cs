namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[] Bytes, OfficeWorkflowConversionEvidence Evidence) SerializeEditableConversion(IOfficeConversionReport report,
        Action<Stream> save, long maximumOutputBytes, CancellationToken token) => SerializeReportedConversion(report, () => {
            using var output = new OfficeWorkflowBoundedMemoryStream(maximumOutputBytes);
            save(output);
            return output.ToArray();
        }, token);

    private static (byte[] Bytes, OfficeWorkflowConversionEvidence Evidence) SerializeReportedConversion(IOfficeConversionReport report,
        Func<byte[]> serialize, CancellationToken token) {
        var evidence = new OfficeWorkflowConversionEvidence(report);
        try {
            byte[] bytes = serialize();
            token.ThrowIfCancellationRequested();
            return (bytes, evidence);
        } catch (OperationCanceledException exception) when (token.IsCancellationRequested) {
            throw new WorkflowConversionCancellationException(exception, evidence);
        } catch (Exception exception) when (exception is not OperationCanceledException and not OutOfMemoryException and not StackOverflowException) {
            throw new WorkflowConversionFailureException(exception, evidence);
        }
    }
}
