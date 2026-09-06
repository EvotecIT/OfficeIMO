using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static OperationArtifact SignPdf(ValidatedRequest request, CancellationToken token) {
        byte[] input = ReadInput(request.InputPath, request.Limits, token);
        PdfDocument source = PdfDocument.Load(input, request.PdfLoadOptions);
        var before = CreateHealthSnapshot(input, request.PdfLoadOptions, token);
        var previous = source.Security.ValidateSignatures();
        var options = request.OutputSignatureOptions!;
        options.MaxInputBytes = Math.Min(options.MaxInputBytes, request.Limits.MaximumInputBytes);
        using var cancellation = CancellationTokenSource.CreateLinkedTokenSource(token, options.CancellationToken);
        options.CancellationToken = cancellation.Token;
        cancellation.Token.ThrowIfCancellationRequested();
        byte[] output = source.Security.SignExternal(request.OutputSigner!, options).Pdf;
        cancellation.Token.ThrowIfCancellationRequested();
        if (output.LongLength > request.Limits.MaximumOutputBytes)
            throw new InvalidOperationException("The signed PDF exceeds the configured output limit.");
        var document = PdfDocument.Load(output, request.OutputPdfLoadOptions);
        var signatures = document.Security.ValidateSignatures(request.OutputSignatureValidator!);
        cancellation.Token.ThrowIfCancellationRequested();
        var after = CreateHealthSnapshot(output, request.OutputPdfLoadOptions, cancellation.Token);
        bool verified = after.CanRead && after.PageCount == before.PageCount &&
            signatures.SignatureCount == previous.SignatureCount + 1 && signatures.IsStructurallyValid &&
            signatures.MathematicalSignaturesVerified && signatures.DigestVerified &&
            signatures.Signatures.Count(signature => signature.Signature.FieldName == options.FieldName) == 1;
        const string summary = "Signed PDF copy created; signature math and document digests verified. Certificate trust is reported separately.";
        return new OperationArtifact(output, summary, new PdfHealthReport(request.Operation, before, after, summary, verified), signatures);
    }
}
