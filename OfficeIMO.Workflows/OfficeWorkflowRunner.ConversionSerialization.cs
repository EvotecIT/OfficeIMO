namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[] Bytes, OfficeWorkflowConversionEvidence Evidence) SerializeEditableConversion(OfficeConversionReport report,
        Action<Stream> save, long maximumOutputBytes, CancellationToken token) => SerializeReportedConversion(report, () => {
            using var output = new OfficeWorkflowBoundedMemoryStream(maximumOutputBytes);
            save(output);
            return output.ToArray();
        }, token);

    private static (byte[] Bytes, OfficeWorkflowConversionEvidence Evidence) SerializeReportedConversion(OfficeConversionReport report,
        Func<byte[]> serialize, CancellationToken token, IReadOnlyDictionary<string, string>? facts = null) {
        var evidence = new OfficeWorkflowConversionEvidence(report, facts);
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
