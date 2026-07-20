using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>
    /// Asynchronously reads a bounded set of files and captures a success or error outcome for every path.
    /// </summary>
    /// <remarks>
    /// Results retain input order. <paramref name="onCompleted"/> is invoked as individual reads finish and
    /// may run concurrently on worker threads. Cancellation is never converted into a failed outcome.
    /// </remarks>
    public Task<IReadOnlyList<ReaderDocumentReadOutcome>> ReadDocumentsDetailedAsync(
        IEnumerable<string> paths,
        ReaderOptions? options = null,
        ReaderBatchOptions? batchOptions = null,
        Action<ReaderDocumentReadOutcome>? onCompleted = null,
        CancellationToken cancellationToken = default) {
        return ReaderBatchExecutor.ExecuteAsync(
            paths,
            batchOptions,
            MaxConcurrentReads,
            MaxConcurrentReads,
            (index, path, token) => ReadDocumentOutcomeAsync(index, path, options, token),
            onCompleted,
            cancellationToken);
    }

    /// <summary>
    /// Asynchronously reads a bounded set of files and reports each resilient outcome without retaining completed
    /// document results for the lifetime of the whole batch.
    /// </summary>
    /// <remarks>
    /// <paramref name="onCompleted"/> is invoked as individual reads finish and may run concurrently on worker
    /// threads. Callers that need an input-ordered materialized result should use
    /// <see cref="ReadDocumentsDetailedAsync"/> instead. Cancellation is never converted into a failed outcome.
    /// </remarks>
    public Task ReadDocumentsAsCompletedAsync(
        IEnumerable<string> paths,
        Action<ReaderDocumentReadOutcome> onCompleted,
        ReaderOptions? options = null,
        ReaderBatchOptions? batchOptions = null,
        CancellationToken cancellationToken = default) {
        return ReaderBatchExecutor.ExecuteAsCompletedAsync(
            paths,
            batchOptions,
            MaxConcurrentReads,
            MaxConcurrentReads,
            (index, path, token) => ReadDocumentOutcomeAsync(index, path, options, token),
            onCompleted,
            cancellationToken);
    }

    private async Task<ReaderDocumentReadOutcome> ReadDocumentOutcomeAsync(
        int index,
        string path,
        ReaderOptions? options,
        CancellationToken cancellationToken) {
        try {
            OfficeDocumentReadResult document = await ReadDocumentAsync(path, options, cancellationToken)
                .ConfigureAwait(false);
            return new ReaderDocumentReadOutcome(index, path, document, error: null);
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception exception) {
            return new ReaderDocumentReadOutcome(index, path, document: null, exception);
        }
    }
}
