using OfficeIMO.Internal;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal async Task<OfficeWorkflowResult> ExtractAsync(IReadOnlyList<int> pageNumbers, string outputPath,
        CancellationToken token, IProgress<PdfWorkspaceProgress>? progress = null,
        OfficeWorkflowStreamOutput? outputStream = null, IOfficeWorkflowPublicationGuard? publicationGuard = null) {
        ThrowIfDisposed();
        if (!CanExtractPages) throw new InvalidOperationException("This document cannot safely extract pages.");
        string destination = OfficeStorageIdentity.Normalize(outputPath);
        if (PathsEqual(destination, Path))
            throw new InvalidOperationException("Extracted pages must be saved to a different file than the open document.");
        if (_storage.UsesProviderPublication(destination) && outputStream is null)
            throw new InvalidOperationException("Provider extraction requires a confirmed output and a recovery store.");
        byte[] snapshot = CopyBytes();
        long revision = Revision;
        string source = Path;
        await VerifyOutputDestinationAsync(destination, token).ConfigureAwait(false);
        var request = new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.ExtractPages,
            InputPath = "urn:officeimo:workspace:" + Guid.NewGuid().ToString("N"),
            InputStream = new("workspace.pdf", _ => Task.FromResult<Stream>(new MemoryStream(snapshot, writable: false))),
            PageNumbers = pageNumbers.ToArray(), PdfPassword = _readOptions.Password,
            OutputPath = destination, OutputStream = outputStream,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
            PublicationGuard = new WorkspaceOutputPublicationGuard(this, source, revision, publicationGuard)
        };
        return await RunNonDetachableCpuWorkAsync(() => new OfficeWorkflowRunner()
            .RunAsync(request, new WorkspaceOutputProgress(progress), token).GetAwaiter().GetResult(), token).ConfigureAwait(false);
    }
}
