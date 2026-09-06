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
            PublicationGuard = new SplitPublicationGuard(this, source, revision, publicationGuard)
        };
        var forwarded = new SplitProgress(progress);
        return await RunNonDetachableCpuWorkAsync(() => new OfficeWorkflowRunner()
            .SplitPdfAsync(request, forwarded, token).GetAwaiter().GetResult(), token).ConfigureAwait(false);
    }

    private sealed class SplitProgress(IProgress<PdfWorkspaceProgress>? progress) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) => progress?.Report(new(value.Message, value.Fraction));
    }

    private sealed class SplitPublicationGuard(PdfWorkspace workspace, string source, long revision,
        IOfficeWorkflowPublicationGuard? host) : IOfficeWorkflowPublicationGuard {
        public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            if (!IsCurrentAndSeparate(path, isDirectory)) return false;
            if (host is not null && !await host.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false)) return false;
            if (workspace._canPublishOutput is not null && !await workspace._canPublishOutput(path, token).ConfigureAwait(false))
                throw new IOException("The output is already owned by an open document. Choose a different destination.");
            token.ThrowIfCancellationRequested();
            return IsCurrentAndSeparate(path, isDirectory);
        }
        private bool IsCurrentAndSeparate(string path, bool isDirectory) {
            workspace.ThrowIfDisposed();
            if (workspace.Revision != revision || !PathsEqual(source, workspace.Path))
                throw new IOException("The document changed while preparing split outputs.");
            if (OfficeStorageIdentity.AreEquivalent(source, path)) return false;
            if (isDirectory && OfficeStorageIdentity.GetLocalPath(source) is { } localSource &&
                OfficeStorageIdentity.GetLocalPath(path) is { } localOutput && OfficePathIdentity.IsSameOrDescendant(localSource, localOutput)) return false;
            workspace._storage.EnsureWritableLocation(path);
            return true;
        }
    }
}
