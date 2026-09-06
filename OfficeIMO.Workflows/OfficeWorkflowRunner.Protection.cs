using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static OperationArtifact ChangeProtection(ValidatedRequest request, CancellationToken token) {
        byte[] input = ReadInput(request.InputPath, request.Limits, token);
        var before = CreateHealthSnapshot(input, request.PdfLoadOptions, token);
        token.ThrowIfCancellationRequested();
        PdfSecurityMutationResult mutation;
        if (request.Operation == OfficeWorkflowOperation.RemovePdfProtection) {
            mutation = PdfSecurityEditor.Decrypt(input, request.PdfOwnerPassword ?? string.Empty,
                request.PdfLoadOptions, request.Limits.MaximumOutputBytes, token);
        } else if (before.HasEncryption) {
            mutation = PdfSecurityEditor.Reencrypt(input, request.PdfOwnerPassword ?? string.Empty, request.OutputEncryption!,
                request.PdfLoadOptions, request.Limits.MaximumOutputBytes, token);
        } else {
            mutation = PdfSecurityEditor.Encrypt(input, request.OutputEncryption!, request.PdfLoadOptions,
                request.Limits.MaximumOutputBytes, token);
        }
        token.ThrowIfCancellationRequested();
        byte[] output = mutation.Pdf;
        var after = CreateHealthSnapshot(output, request.OutputPdfLoadOptions, token);
        bool expectedEncryption = request.Operation == OfficeWorkflowOperation.ProtectPdf;
        bool verified = mutation.PreservationReport.IsPreserved && after.CanRead && after.PageCount == before.PageCount &&
            after.HasEncryption == expectedEncryption;
        if (expectedEncryption) {
            var settings = request.OutputEncryption!;
            var security = mutation.OutputSecurity;
            verified &= security.EncryptionPermissions == settings.Permissions && (security.EncryptMetadata ?? true) == settings.EncryptMetadata;
            var userOptions = PdfLoadOptions.WithPassword(request.OutputPdfLoadOptions, settings.UserPassword);
            var userSecurity = PdfSyntax.ReadDocumentSecurityInfo(output, userOptions, cancellationToken: token);
            verified &= userSecurity.PasswordAuthenticationRole != PdfPasswordAuthenticationRole.None;
        }
        string summary = expectedEncryption ? "Protected PDF copy created and verified." : "Unencrypted PDF copy created and verified.";
        var report = new PdfHealthReport(request.Operation, before, after, summary, verified);
        return new OperationArtifact(output, summary, report);
    }
}
