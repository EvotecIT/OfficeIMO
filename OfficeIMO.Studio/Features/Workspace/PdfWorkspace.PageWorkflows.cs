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
        ThrowIfDisposed();
        if (!CanImportPages) throw new InvalidOperationException("This document cannot safely import pages.");
        if (sourcePaths is null || sourcePaths.Count == 0) {
            throw new ArgumentException("Choose at least one PDF to import.", nameof(sourcePaths));
        }

        string[] sources = sourcePaths.Select(ValidateSourcePdfPath).ToArray();
        var snapshots = new List<byte[]>(sources.Length);
        long remainingBytes = OfficeIMO.Studio.Infrastructure.StudioStorageAccess.MaximumDocumentBytes;
        foreach (string source in sources) {
            cancellationToken.ThrowIfCancellationRequested();
            if (remainingBytes == 0) throw new IOException("The selected PDFs exceed the 512 MiB import limit. Import fewer documents at a time.");
            var snapshot = await _storage.ReadSnapshotAsync(source, cancellationToken, remainingBytes).ConfigureAwait(false);
            snapshots.Add(snapshot.Bytes);
            remainingBytes -= snapshot.Bytes.LongLength;
        }
        int importedPageCount = 0;
        string description = sources.Length == 1
            ? $"Imported pages from {_storage.Describe(sources[0]).Name}"
            : $"Imported pages from {sources.Length} PDF documents";

        await MutateAsync(
            PdfWorkspaceOperationKind.Import,
            description,
            Array.Empty<int>(),
            document => {
                int insertionPage = insertBeforePageNumber;
                foreach (byte[] sourceBytes in snapshots) {
                    cancellationToken.ThrowIfCancellationRequested();
                    PdfDocument source = PdfDocument.Load(sourceBytes);
                    int sourcePageCount = source.Inspect().PageCount;
                    document = document.Pages.Insert(insertionPage, source);
                    insertionPage += sourcePageCount;
                    importedPageCount += sourcePageCount;
                }
                return document;
            },
            cancellationToken,
            progress).ConfigureAwait(false);

        return importedPageCount;
    }

    internal async Task ExtractAsync(
        IReadOnlyList<int> pageNumbers,
        string outputPath,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        ThrowIfDisposed();
        if (!CanExtractPages) throw new InvalidOperationException("This document cannot safely extract pages.");
        string destination = OfficeIMO.Internal.OfficeStorageIdentity.Normalize(outputPath);
        if (PathsEqual(destination, Path)) {
            throw new InvalidOperationException("Extracted pages must be saved to a different file than the open document.");
        }

        progress?.Report(new PdfWorkspaceProgress("Extracting pages", 0.1D));
        PdfDocument extracted = await RunCancellableCpuWorkAsync(
            () => CreateDocumentSnapshot().Pages.Extract(pageNumbers.ToArray()),
            cancellationToken).ConfigureAwait(false);
        progress?.Report(new PdfWorkspaceProgress("Saving extracted PDF", 0.7D));
        await WriteWorkspaceOutputAsync(destination,
            (stream, token) => extracted.SaveAsync(stream, token), cancellationToken).ConfigureAwait(false);
        progress?.Report(new PdfWorkspaceProgress("Extract complete", 1D));
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
