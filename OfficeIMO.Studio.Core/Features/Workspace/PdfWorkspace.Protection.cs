using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal Task<OfficeWorkflowResult> SaveProtectedCopyAsync(string destinationPath,
        PdfStandardEncryptionOptions encryption, string? currentOwnerPassword, CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null, OfficeWorkflowStreamOutput? outputStream = null,
        IOfficeWorkflowPublicationGuard? publicationGuard = null) {
        ArgumentNullException.ThrowIfNull(encryption);
        return ChangeProtectionAsync(destinationPath, encryption.Clone(), currentOwnerPassword, cancellationToken,
            progress, outputStream, publicationGuard);
    }

    internal Task<OfficeWorkflowResult> SaveDecryptedCopyAsync(string destinationPath,
        string? ownerPassword, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null,
        OfficeWorkflowStreamOutput? outputStream = null, IOfficeWorkflowPublicationGuard? publicationGuard = null) =>
        ChangeProtectionAsync(destinationPath, null, ownerPassword, cancellationToken, progress, outputStream, publicationGuard);

    private async Task<OfficeWorkflowResult> ChangeProtectionAsync(string destinationPath,
        PdfStandardEncryptionOptions? encryption, string? ownerPassword, CancellationToken token,
        IProgress<PdfWorkspaceProgress>? progress, OfficeWorkflowStreamOutput? outputStream,
        IOfficeWorkflowPublicationGuard? publicationGuard) {
        ThrowIfDisposed();
        if (encryption is null && !IsEncrypted) throw new InvalidOperationException("This PDF is not password protected.");
        if (!CanChangeEncryption(ownerPassword))
            throw new InvalidOperationException("This document's signature, certification, usage-rights, or authorization policy prevents changing password protection.");
        string destination = ValidateExportDestination(destinationPath);
        if (_storage.UsesProviderPublication(destination) && outputStream is null)
            throw new InvalidOperationException("Provider protection output requires a confirmed destination and a recovery store.");
        byte[] snapshot = CopyBytes();
        long revision = Revision;
        string source = Path;
        string? effectiveOwner = _documentInfo.Security.HasOwnerAuthorization ? _readOptions.Password : ownerPassword;
        string? inputPassword = _readOptions.Password;
        await VerifyOutputDestinationAsync(destination, token).ConfigureAwait(false);
        var request = new OfficeWorkflowRequest {
            Operation = encryption is null ? OfficeWorkflowOperation.RemovePdfProtection : OfficeWorkflowOperation.ProtectPdf,
            InputPath = "urn:officeimo:workspace:" + Guid.NewGuid().ToString("N"),
            InputStream = new("workspace.pdf", _ => Task.FromResult<Stream>(new MemoryStream(snapshot, writable: false))),
            PdfPassword = inputPassword, PdfOwnerPassword = effectiveOwner, OutputEncryption = encryption,
            OutputPath = destination, OutputStream = outputStream, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
            PublicationGuard = new WorkspaceOutputPublicationGuard(this, source, revision, publicationGuard)
        };
        return await RunNonDetachableCpuWorkAsync(() => new OfficeWorkflowRunner()
            .RunAsync(request, new WorkspaceOutputProgress(progress), token).GetAwaiter().GetResult(), token).ConfigureAwait(false);
    }
}
