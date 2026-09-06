using OfficeIMO.Internal;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal async Task<PdfSplitWorkflowResult> SplitAsync(string destination, int pagesPerDocument,
        CancellationToken token, IProgress<PdfWorkspaceProgress>? progress = null,
        OfficeWorkflowDirectoryOutput? directoryOutput = null, IOfficeWorkflowPublicationGuard? publicationGuard = null) {
        ThrowIfDisposed();
        if (!CanExtractPages) throw new InvalidOperationException("This document cannot safely split pages.");
        byte[] snapshot = CopyBytes();
        long revision = Revision;
        string source = Path;
        var request = new PdfSplitWorkflowRequest {
            // This immutable input is the edited workspace, which may differ from its backing file.
            InputPath = "urn:officeimo:workspace:" + Guid.NewGuid().ToString("N"),
            InputStream = new("workspace.pdf", _ => Task.FromResult<Stream>(new MemoryStream(snapshot, writable: false))),
            PdfPassword = _readOptions.Password, OutputDirectory = destination, DirectoryOutput = directoryOutput,
            PagesPerDocument = pagesPerDocument,
            ConflictPolicy = directoryOutput is null ? OfficeWorkflowConflictPolicy.Rename : OfficeWorkflowConflictPolicy.Replace,
            PublicationGuard = new WorkspaceOutputPublicationGuard(this, source, revision, publicationGuard)
        };
        var forwarded = new WorkspaceOutputProgress(progress);
        return await RunNonDetachableCpuWorkAsync(() => new OfficeWorkflowRunner()
            .SplitPdfAsync(request, forwarded, token).GetAwaiter().GetResult(), token).ConfigureAwait(false);
    }

}
