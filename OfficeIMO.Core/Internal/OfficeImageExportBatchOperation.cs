using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

/// <summary>
/// Runs bounded batch rendering and persistence without a fluent builder. Payload-releasing
/// saves and the one-item asynchronous bridge share the same execution and policy checks.
/// </summary>
internal sealed class OfficeImageExportBatchOperation<TOptions> where TOptions : OfficeImageExportOptions {
    private readonly TOptions _options;
    private readonly OfficeImageExportFormat _format;
    private readonly OfficeImageExportFileConflictPolicy _conflictPolicy;
    private readonly Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> _export;
    private readonly Func<OfficeImageExportFormat, TOptions, CancellationToken, Task<IReadOnlyList<OfficeImageExportResult>>>? _exportAsync;
    private readonly Action<OfficeImageExportFormat, TOptions, OfficeImageExportConsumer, CancellationToken>? _exportEach;
    private readonly Func<OfficeImageExportFormat, TOptions, OfficeImageExportAsyncConsumer, CancellationToken, Task>? _exportEachAsync;

    internal OfficeImageExportBatchOperation(
        TOptions options,
        OfficeImageExportFormat format,
        OfficeImageExportFileConflictPolicy conflictPolicy,
        Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> export,
        Func<OfficeImageExportFormat, TOptions, CancellationToken, Task<IReadOnlyList<OfficeImageExportResult>>>? exportAsync = null,
        Action<OfficeImageExportFormat, TOptions, OfficeImageExportConsumer, CancellationToken>? exportEach = null,
        Func<OfficeImageExportFormat, TOptions, OfficeImageExportAsyncConsumer, CancellationToken, Task>? exportEachAsync = null) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        if (!Enum.IsDefined(typeof(OfficeImageExportFormat), format)) throw new ArgumentOutOfRangeException(nameof(format));
        _format = format;
        _conflictPolicy = conflictPolicy;
        _export = export ?? throw new ArgumentNullException(nameof(export));
        _exportAsync = exportAsync;
        _exportEach = exportEach;
        _exportEachAsync = exportEachAsync;
    }

    /// <summary>Exports all selected images and materializes their encoded payloads.</summary>
    internal IReadOnlyList<OfficeImageExportResult> Export() {
        var results = new List<OfficeImageExportResult>();
        ExportEach(results.Add);
        FinalizeSequence(results);
        return results.AsReadOnly();
    }

    /// <summary>Asynchronously exports all selected images and materializes their encoded payloads.</summary>
    internal async Task<IReadOnlyList<OfficeImageExportResult>> ExportAsync(
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        await ExportEachAsync(
            (result, _) => {
                results.Add(result);
                return Task.CompletedTask;
            },
            cancellationToken).ConfigureAwait(false);
        FinalizeSequence(results);
        return results.AsReadOnly();
    }

    /// <summary>Streams results to a consumer so callers can release each payload immediately.</summary>
    internal void ExportEach(OfficeImageExportConsumer consumer, CancellationToken cancellationToken = default) {
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        cancellationToken.ThrowIfCancellationRequested();
        TOptions effective = _options.CreateEffectiveImageExportOptions<TOptions>();
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            effective.RenderTimeout,
            cancellationToken);
        try {
            var tracker = new OfficeImageExportBatchTracker(effective);
            int completed = 0;
            effective.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Rendering, 0));

            void Accept(OfficeImageExportResult result) {
                execution.Token.ThrowIfCancellationRequested();
                effective.EnsureAccepted(result);
                tracker.Add(result);
                int? sequenceCount = result.SequenceIndex == completed ? result.SequenceCount : null;
                OfficeImageExportResult sequenced = result.WithSequence(completed, sequenceCount);
                consumer(sequenced);
                completed++;
                effective.Progress?.Report(new OfficeImageExportProgress(
                    OfficeImageExportProgressStage.Completed,
                    completed,
                    name: sequenced.Name));
            }

            if (_exportEach != null) {
                _exportEach(_format, effective, Accept, execution.Token);
                execution.ThrowIfCancellationRequested();
                return;
            }

            foreach (OfficeImageExportResult result in _export(_format, effective)) Accept(result);
            execution.ThrowIfCancellationRequested();
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    /// <summary>Streams results asynchronously when the adapter owns a genuine asynchronous path.</summary>
    internal async Task ExportEachAsync(
        OfficeImageExportAsyncConsumer consumer,
        CancellationToken cancellationToken = default) {
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        cancellationToken.ThrowIfCancellationRequested();
        TOptions effective = _options.CreateEffectiveImageExportOptions<TOptions>();
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            effective.RenderTimeout,
            cancellationToken);
        CancellationToken operationCancellationToken = execution.Token;
        var tracker = new OfficeImageExportBatchTracker(effective);
        int completed = 0;
        effective.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Rendering, 0));

        async Task AcceptAsync(OfficeImageExportResult result, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            effective.EnsureAccepted(result);
            tracker.Add(result);
            int? sequenceCount = result.SequenceIndex == completed ? result.SequenceCount : null;
            OfficeImageExportResult sequenced = result.WithSequence(completed, sequenceCount);
            await consumer(sequenced, token).ConfigureAwait(false);
            completed++;
            effective.Progress?.Report(new OfficeImageExportProgress(
                OfficeImageExportProgressStage.Completed,
                completed,
                name: sequenced.Name));
        }

        try {
            if (_exportEachAsync != null) {
                await _exportEachAsync(_format, effective, AcceptAsync, operationCancellationToken).ConfigureAwait(false);
                execution.ThrowIfCancellationRequested();
                return;
            }

            if (_exportAsync != null) {
                IReadOnlyList<OfficeImageExportResult> asyncResults =
                    await _exportAsync(_format, effective, operationCancellationToken).ConfigureAwait(false);
                foreach (OfficeImageExportResult result in asyncResults) await AcceptAsync(result, operationCancellationToken).ConfigureAwait(false);
                execution.ThrowIfCancellationRequested();
                return;
            }

            if (_exportEach != null) {
                using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(operationCancellationToken);
                using var buffer = new OfficeImageExportAsyncBuffer();
                Task producer = Task.Run(() => {
                    try {
                        _exportEach(
                            _format,
                            effective,
                            result => buffer.Add(result, linkedCancellation.Token),
                            linkedCancellation.Token);
                        buffer.Complete();
                    } catch (Exception exception) {
                        buffer.Complete(exception);
                    }
                });
                try {
                    while (await buffer.ReadAsync(operationCancellationToken).ConfigureAwait(false) is { } result) {
                        await AcceptAsync(result, operationCancellationToken).ConfigureAwait(false);
                    }
                } finally {
                    linkedCancellation.Cancel();
                    await producer.ConfigureAwait(false);
                }
                execution.ThrowIfCancellationRequested();
                return;
            }

            foreach (OfficeImageExportResult result in _export(_format, effective)) {
                await AcceptAsync(result, operationCancellationToken).ConfigureAwait(false);
            }
            execution.ThrowIfCancellationRequested();
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    /// <summary>Saves all selected images and returns payload-bearing results with their normalized paths.</summary>
    internal IReadOnlyList<OfficeImageExportResult> Save(string folderPath) {
        string fullFolder = PrepareFolder(folderPath);
        var saved = new List<OfficeImageExportResult>();
        var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int index = 0;
        ExportEach(result => {
            string path = ResolveBatchPath(fullFolder, result, index++, usedFileNames);
            _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, index - 1, name: result.Name, destinationPath: path));
            string savedPath = result.WriteToFile(path, _conflictPolicy);
            saved.Add(result.WithSavedPath(savedPath));
        });
        FinalizeSequence(saved);
        return saved.AsReadOnly();
    }

    /// <summary>
    /// Saves all selected images and returns only paths, metadata, and diagnostics so encoded payloads can be released.
    /// </summary>
    internal OfficeImageExportBatchSaveResult SaveFiles(string folderPath, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        string fullFolder = PrepareFolder(folderPath);
        var files = new List<OfficeImageExportSavedFile>();
        var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int index = 0;
        ExportEach(result => {
            string path = ResolveBatchPath(fullFolder, result, index++, usedFileNames);
            _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, index - 1, name: result.Name, destinationPath: path));
            string savedPath = result.WriteToFile(path, _conflictPolicy);
            files.Add(new OfficeImageExportSavedFile(result, savedPath));
        }, cancellationToken);
        FinalizeSequence(files);
        return new OfficeImageExportBatchSaveResult(files);
    }

    /// <summary>Asynchronously saves all selected images and returns payload-bearing results with normalized paths.</summary>
    internal async Task<IReadOnlyList<OfficeImageExportResult>> SaveAsync(
        string folderPath,
        CancellationToken cancellationToken = default) {
        string fullFolder = PrepareFolder(folderPath);
        var saved = new List<OfficeImageExportResult>();
        var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int index = 0;
        await ExportEachAsync(async (result, token) => {
            string path = ResolveBatchPath(fullFolder, result, index++, usedFileNames);
            _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, index - 1, name: result.Name, destinationPath: path));
            string savedPath = await result.WriteToFileAsync(path, _conflictPolicy, token).ConfigureAwait(false);
            saved.Add(result.WithSavedPath(savedPath));
        }, cancellationToken).ConfigureAwait(false);
        FinalizeSequence(saved);
        return saved.AsReadOnly();
    }

    /// <summary>Asynchronously saves images without retaining encoded payloads in the returned result.</summary>
    internal async Task<OfficeImageExportBatchSaveResult> SaveFilesAsync(
        string folderPath,
        CancellationToken cancellationToken = default) {
        string fullFolder = PrepareFolder(folderPath);
        var files = new List<OfficeImageExportSavedFile>();
        var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int index = 0;
        await ExportEachAsync(async (result, token) => {
            string path = ResolveBatchPath(fullFolder, result, index++, usedFileNames);
            _options.Progress?.Report(new OfficeImageExportProgress(OfficeImageExportProgressStage.Saving, index - 1, name: result.Name, destinationPath: path));
            string savedPath = await result.WriteToFileAsync(path, _conflictPolicy, token).ConfigureAwait(false);
            files.Add(new OfficeImageExportSavedFile(result, savedPath));
        }, cancellationToken).ConfigureAwait(false);
        FinalizeSequence(files);
        return new OfficeImageExportBatchSaveResult(files);
    }

    private string ResolveBatchPath(
        string folder,
        OfficeImageExportResult result,
        int index,
        ISet<string> usedFileNames) {
        string name = string.IsNullOrWhiteSpace(result.Name)
            ? "image-" + (index + 1).ToString(CultureInfo.InvariantCulture)
            : result.Name!;
        string fileName = GetUniqueFileName(OfficePortableFileName.SanitizeBaseName(name), _format.GetFileExtension(), usedFileNames);
        return OfficeImageExportPath.NormalizeFile(Path.Combine(folder, fileName), _format);
    }

    private static void FinalizeSequence(List<OfficeImageExportResult> results) {
        for (int index = 0; index < results.Count; index++) {
            results[index] = results[index].WithSequence(index, results.Count);
        }
    }

    private static void FinalizeSequence(List<OfficeImageExportSavedFile> files) {
        for (int index = 0; index < files.Count; index++) {
            files[index] = files[index].WithSequence(index, files.Count);
        }
    }

    private static string PrepareFolder(string folderPath) {
        if (string.IsNullOrWhiteSpace(folderPath)) throw new ArgumentException("Output folder cannot be null or whitespace.", nameof(folderPath));
        string fullFolder = Path.GetFullPath(folderPath);
        Directory.CreateDirectory(fullFolder);
        return fullFolder;
    }

    private static string GetUniqueFileName(string baseName, string extension, ISet<string> usedFileNames) {
        if (string.IsNullOrWhiteSpace(baseName)) baseName = "image";
        string candidate = baseName + extension;
        if (usedFileNames.Add(candidate)) return candidate;

        int suffix = 2;
        do {
            candidate = baseName + "-" + suffix.ToString(CultureInfo.InvariantCulture) + extension;
            suffix++;
        } while (!usedFileNames.Add(candidate));
        return candidate;
    }

}
