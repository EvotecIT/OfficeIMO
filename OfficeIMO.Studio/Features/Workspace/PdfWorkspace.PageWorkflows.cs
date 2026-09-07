using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal Task<int> ImportAsync(
        string sourcePath,
        int insertBeforePageNumber,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        ImportAsync(new[] { sourcePath }, insertBeforePageNumber, cancellationToken, progress);

    internal async Task<int> ImportAsync(
        IReadOnlyList<string> sourcePaths,
        int insertBeforePageNumber,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        var preparation = await PrepareImportAsync(sourcePaths, cancellationToken).ConfigureAwait(false)
            ?? throw new OperationCanceledException(cancellationToken);
        var selections = preparation.Sources.Select((source, index) =>
            new PdfImportSelection(index, Enumerable.Range(1, source.PageCount).ToArray())).ToArray();
        return await ApplyImportAsync(preparation, selections, insertBeforePageNumber, cancellationToken, progress).ConfigureAwait(false);
    }
    private string ValidateSourcePdfPath(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("PDF source paths cannot be empty.", nameof(path));
        string fullPath = OfficeIMO.Internal.OfficeStorageIdentity.Normalize(path);
        string name = _storage.Describe(fullPath).Name;
        if (!string.Equals(System.IO.Path.GetExtension(name), ".pdf", StringComparison.OrdinalIgnoreCase)) {
            throw new NotSupportedException($"Only PDF documents can be imported: {name}");
        }
        return fullPath;
    }

    internal async Task<T> RunCancellableCpuWorkAsync<T>(
        Func<T> operation,
        CancellationToken cancellationToken) =>
        await RunCpuWorkAsync(operation, cancellationToken, detachOnCancellation: true).ConfigureAwait(false);

    internal async Task<T> RunNonDetachableCpuWorkAsync<T>(
        Func<T> operation,
        CancellationToken cancellationToken) =>
        await RunCpuWorkAsync(operation, cancellationToken, detachOnCancellation: false).ConfigureAwait(false);

    private async Task<T> RunCpuWorkAsync<T>(
        Func<T> operation,
        CancellationToken cancellationToken,
        bool detachOnCancellation) {
        ArgumentNullException.ThrowIfNull(operation);
        cancellationToken.ThrowIfCancellationRequested();
        await ApplicationCpuWorkGate.WaitAsync(cancellationToken).ConfigureAwait(false);

        Task<T> worker;
        try {
            ThrowIfDisposed();
            cancellationToken.ThrowIfCancellationRequested();
            worker = Task.Run(operation, CancellationToken.None);
            lock (_cpuWorkSync) {
                _activeCpuWorker = worker;
            }
        } catch {
            ApplicationCpuWorkGate.Release();
            throw;
        }

        _ = CompleteCpuWorkAsync(worker);
        return detachOnCancellation
            ? await worker.WaitAsync(cancellationToken).ConfigureAwait(false)
            : await worker.ConfigureAwait(false);
    }

    private async Task CompleteCpuWorkAsync(Task worker) {
        try {
            await worker.ConfigureAwait(false);
        } catch {
            // The caller observes failures while attached. A cancelled caller has already been notified.
        } finally {
            lock (_cpuWorkSync) {
                if (ReferenceEquals(_activeCpuWorker, worker)) _activeCpuWorker = null;
            }
            ApplicationCpuWorkGate.Release();
        }
    }

    private static bool PathsEqual(string left, string right) => OfficeIMO.Internal.OfficeStorageIdentity.AreEquivalent(left, right);
}
