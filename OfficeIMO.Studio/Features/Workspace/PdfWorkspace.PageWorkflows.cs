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
        int importedPageCount = 0;
        string description = sources.Length == 1
            ? $"Imported pages from {System.IO.Path.GetFileName(sources[0])}"
            : $"Imported pages from {sources.Length} PDF documents";

        await MutateAsync(
            PdfWorkspaceOperationKind.Import,
            description,
            Array.Empty<int>(),
            document => {
                int insertionPage = insertBeforePageNumber;
                foreach (string sourcePath in sources) {
                    cancellationToken.ThrowIfCancellationRequested();
                    PdfDocument source = PdfDocument.Load(sourcePath);
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
        string destination = System.IO.Path.GetFullPath(outputPath);
        if (PathsEqual(destination, Path)) {
            throw new InvalidOperationException("Extracted pages must be saved to a different file than the open document.");
        }

        progress?.Report(new PdfWorkspaceProgress("Extracting pages", 0.1D));
        PdfDocument extracted = await RunCancellableCpuWorkAsync(
            () => CreateDocumentSnapshot().Pages.Extract(pageNumbers.ToArray()),
            cancellationToken).ConfigureAwait(false);
        progress?.Report(new PdfWorkspaceProgress("Saving extracted PDF", 0.7D));
        await extracted.SaveAsync(destination, cancellationToken).ConfigureAwait(false);
        progress?.Report(new PdfWorkspaceProgress("Extract complete", 1D));
    }

    internal async Task<IReadOnlyList<string>> SplitAsync(
        string outputDirectory,
        int pagesPerDocument,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        ThrowIfDisposed();
        if (!CanExtractPages) throw new InvalidOperationException("This document cannot safely split pages.");
        if (pagesPerDocument < 1) throw new ArgumentOutOfRangeException(nameof(pagesPerDocument));

        string destinationRoot = System.IO.Path.GetFullPath(outputDirectory);
        bool destinationExisted = Directory.Exists(destinationRoot);
        Directory.CreateDirectory(destinationRoot);
        IReadOnlyList<PdfDocument> outputs = await RunCancellableCpuWorkAsync(
            () => CreateDocumentSnapshot().Pages.Split(pagesPerDocument),
            cancellationToken).ConfigureAwait(false);

        string stem = System.IO.Path.GetFileNameWithoutExtension(Path);
        string[] destinations = Enumerable.Range(1, outputs.Count)
            .Select(index => System.IO.Path.Combine(destinationRoot, $"{stem}-part-{index:D3}.pdf"))
            .ToArray();
        string[] collisions = destinations
            .Where(path => File.Exists(path) || Directory.Exists(path))
            .Select(path => System.IO.Path.GetFileName(path) ?? path)
            .ToArray();
        if (collisions.Length > 0) {
            throw new IOException(
                collisions.Length == 1
                    ? $"The split output {collisions[0]} already exists. Choose another folder or move the existing file."
                    : $"{collisions.Length} split outputs already exist. Choose another folder or move the existing files.");
        }

        string stagingRoot = System.IO.Path.Combine(destinationRoot, $".officeimo-studio-split-{Guid.NewGuid():N}");
        var committed = new List<string>();
        try {
            Directory.CreateDirectory(stagingRoot);
            for (int index = 0; index < outputs.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(new PdfWorkspaceProgress(
                    $"Preparing part {index + 1} of {outputs.Count}",
                    0.1D + (0.7D * index / Math.Max(1, outputs.Count))));
                string stagedPath = System.IO.Path.Combine(stagingRoot, System.IO.Path.GetFileName(destinations[index]));
                await outputs[index].SaveAsync(stagedPath, cancellationToken).ConfigureAwait(false);
            }

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new PdfWorkspaceProgress("Publishing split files", 0.9D));
            for (int index = 0; index < destinations.Length; index++) {
                string stagedPath = System.IO.Path.Combine(stagingRoot, System.IO.Path.GetFileName(destinations[index]));
                File.Move(stagedPath, destinations[index]);
                committed.Add(destinations[index]);
            }

            progress?.Report(new PdfWorkspaceProgress("Split complete", 1D));
            return destinations;
        } catch {
            foreach (string path in committed) {
                if (File.Exists(path)) File.Delete(path);
            }
            throw;
        } finally {
            if (Directory.Exists(stagingRoot)) Directory.Delete(stagingRoot, recursive: true);
            if (!destinationExisted && Directory.Exists(destinationRoot) && !Directory.EnumerateFileSystemEntries(destinationRoot).Any()) {
                Directory.Delete(destinationRoot);
            }
        }
    }

    private static string ValidateSourcePdfPath(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("PDF source paths cannot be empty.", nameof(path));
        string fullPath = System.IO.Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("An imported PDF no longer exists.", fullPath);
        if (!string.Equals(System.IO.Path.GetExtension(fullPath), ".pdf", StringComparison.OrdinalIgnoreCase)) {
            throw new NotSupportedException($"Only PDF documents can be imported: {System.IO.Path.GetFileName(fullPath)}");
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

    private static bool PathsEqual(string left, string right) => string.Equals(
        System.IO.Path.GetFullPath(left),
        System.IO.Path.GetFullPath(right),
        OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
}
