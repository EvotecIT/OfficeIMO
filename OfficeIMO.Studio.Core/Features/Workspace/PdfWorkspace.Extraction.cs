using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Infrastructure;
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

    /// <summary>
    /// Writes pages, in the order given, to a new PDF in a local folder the host owns. It uses the same page extraction
    /// and output publication as the workflow, without the workflow runner and its conversions, so a NativeAOT host
    /// stays small; provider locations go through <see cref="ExtractAsync"/>.
    /// </summary>
    internal async Task ExtractToLocalFileAsync(IReadOnlyList<int> pageNumbers, string outputPath, CancellationToken token) {
        ThrowIfDisposed();
        if (!CanExtractPages) throw new InvalidOperationException("This document cannot safely extract pages.");
        if (pageNumbers.Count == 0) throw new ArgumentException("Choose at least one page.", nameof(pageNumbers));
        string destination = OfficeStorageIdentity.Normalize(outputPath);
        if (PathsEqual(destination, Path))
            throw new InvalidOperationException("Extracted pages must be saved to a different file than the open document.");
        if (_storage.UsesProviderPublication(destination))
            throw new NotSupportedException("Extracting to a provider location requires the workflow output.");
        byte[] snapshot = CopyBytes();
        int[] pages = pageNumbers.ToArray();
        await VerifyOutputDestinationAsync(destination, token).ConfigureAwait(false);
        byte[] bytes = await RunNonDetachableCpuWorkAsync(() => {
            PdfDocument extracted = LoadDocument(snapshot).Pages.Extract(pages, StudioDocumentStorage.MaximumDocumentBytes);
            token.ThrowIfCancellationRequested();
            using var output = new OfficeBoundedMemoryStream(StudioDocumentStorage.MaximumDocumentBytes);
            extracted.SaveAsync(output, token).GetAwaiter().GetResult();
            byte[] saved = output.ToArray();
            if (PdfDocument.Load(saved).Inspect().PageCount != pages.Length)
                throw new InvalidDataException("The saved extraction does not contain the requested number of pages.");
            return saved;
        }, token).ConfigureAwait(false);
        await WriteWorkspaceOutputAsync(destination, (stream, writeToken) => stream.WriteAsync(bytes.AsMemory(), writeToken).AsTask(), token)
            .ConfigureAwait(false);
    }
}
