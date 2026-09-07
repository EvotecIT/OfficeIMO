using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal Task<byte[]?> PreviewSignatureAsync(PdfExternalSignatureOptions options, CancellationToken token) {
        ThrowIfDisposed();
        var document = CreateDocumentSnapshot();
        return RunCancellableCpuWorkAsync<byte[]?>(() => {
            options.CancellationToken = token;
            byte[] prepared = document.Security.PrepareExternalSignature(options).PreparedPdf;
            token.ThrowIfCancellationRequested();
            if (options.VisibleAppearance is not { } appearance) return null;
            var candidate = PdfDocument.Load(prepared, _readOptions);
            var page = candidate.InspectForViewing(cancellationToken: token).Pages[appearance.PageNumber - 1];
            var rendered = candidate.Render.DisplayPage(appearance.PageNumber, new PdfPageDisplayOptions {
                Scale = Math.Min(1D, 700D / Math.Max(page.Width, page.Height)), MaximumOutputBytes = 8 * 1024 * 1024
            }, token);
            if (!rendered.Succeeded || rendered.Bytes is null) throw new InvalidOperationException(string.Join(Environment.NewLine, rendered.Diagnostics));
            return rendered.Bytes;
        }, token);
    }

    internal async Task<OfficeWorkflowResult> SaveSignedCopyAsync(string destinationPath, IPdfExternalSigner signer,
        PdfExternalSignatureOptions options, IPdfSignatureCryptographyProvider verifier, CancellationToken token,
        IProgress<PdfWorkspaceProgress>? progress = null, OfficeWorkflowStreamOutput? outputStream = null,
        IOfficeWorkflowPublicationGuard? publicationGuard = null) {
        ThrowIfDisposed();
        if (!CanSign) throw new InvalidOperationException("This document's security policy prevents signing.");
        string destination = ValidateExportDestination(destinationPath);
        if (_storage.UsesProviderPublication(destination) && outputStream is null)
            throw new InvalidOperationException("Provider signing output requires a confirmed destination and a recovery store.");
        byte[] snapshot = CopyBytes(); long revision = Revision; string source = Path;
        // The fluent owner captures all signature and appearance settings before the destination check can await.
        var request = OfficeWorkflow.SignPdf("urn:officeimo:workspace:" + Guid.NewGuid().ToString("N"), signer, options, verifier)
            .WithPdfPassword(_readOptions.Password).To(destination).OnConflict(OfficeWorkflowConflictPolicy.Replace).Build();
        request.InputStream = new("workspace.pdf", _ => Task.FromResult<Stream>(new MemoryStream(snapshot, false)));
        request.OutputStream = outputStream;
        request.PublicationGuard = new WorkspaceOutputPublicationGuard(this, source, revision, publicationGuard);
        await VerifyOutputDestinationAsync(destination, token).ConfigureAwait(false);
        return await RunNonDetachableCpuWorkAsync(() => new OfficeWorkflowRunner()
            .RunAsync(request, new WorkspaceOutputProgress(progress), token).GetAwaiter().GetResult(), token).ConfigureAwait(false);
    }
}
