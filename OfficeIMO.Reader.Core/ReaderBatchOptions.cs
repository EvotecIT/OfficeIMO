using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

/// <summary>
/// Bounds multi-document asynchronous Reader execution.
/// </summary>
public sealed class ReaderBatchOptions {
    /// <summary>
    /// Maximum number of document reads in flight. Null uses the reader's configured limit.
    /// </summary>
    public int? MaxDegreeOfParallelism { get; set; }

    /// <summary>
    /// Maximum number of input documents accepted by one batch. Default: 500.
    /// </summary>
    public int MaxDocuments { get; set; } = 500;
}

internal static class ReaderBatchExecutor {
    public static async Task<IReadOnlyList<T>> ExecuteAsync<T>(
        IEnumerable<string> paths,
        ReaderBatchOptions? options,
        int defaultMaxDegreeOfParallelism,
        int maxDegreeOfParallelismLimit,
        Func<string, CancellationToken, Task<T>> readAsync,
        CancellationToken cancellationToken) {
        if (paths == null) throw new ArgumentNullException(nameof(paths));
        if (readAsync == null) throw new ArgumentNullException(nameof(readAsync));
        if (defaultMaxDegreeOfParallelism < 1) throw new ArgumentOutOfRangeException(nameof(defaultMaxDegreeOfParallelism));
        if (maxDegreeOfParallelismLimit < 1) throw new ArgumentOutOfRangeException(nameof(maxDegreeOfParallelismLimit));

        int maxDocuments = options?.MaxDocuments ?? 500;
        if (maxDocuments < 1) {
            throw new ArgumentOutOfRangeException(nameof(options), "MaxDocuments must be greater than 0.");
        }

        int requestedDegree = options?.MaxDegreeOfParallelism ?? defaultMaxDegreeOfParallelism;
        if (requestedDegree < 1) {
            throw new ArgumentOutOfRangeException(nameof(options), "MaxDegreeOfParallelism must be greater than 0 when specified.");
        }

        int degree = Math.Min(requestedDegree, maxDegreeOfParallelismLimit);
        var inputs = new List<string>(Math.Min(maxDocuments, 256));
        foreach (string path in paths) {
            cancellationToken.ThrowIfCancellationRequested();
            if (inputs.Count >= maxDocuments) {
                throw new InvalidOperationException($"Batch input exceeds MaxDocuments ({maxDocuments}).");
            }
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Batch paths cannot contain null or empty values.", nameof(paths));
            }

            inputs.Add(path);
        }

        if (inputs.Count == 0) {
            return Array.Empty<T>();
        }

        var results = new T[inputs.Count];
        int nextIndex = -1;
        int workerCount = Math.Min(degree, inputs.Count);
        using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var workers = new Task[workerCount];

        for (int workerIndex = 0; workerIndex < workerCount; workerIndex++) {
            workers[workerIndex] = RunWorkerAsync();
        }

        await Task.WhenAll(workers).ConfigureAwait(false);
        return results;

        async Task RunWorkerAsync() {
            try {
                while (true) {
                    linkedCancellation.Token.ThrowIfCancellationRequested();
                    int index = Interlocked.Increment(ref nextIndex);
                    if (index >= inputs.Count) {
                        return;
                    }

                    results[index] = await readAsync(inputs[index], linkedCancellation.Token).ConfigureAwait(false);
                }
            } catch {
                linkedCancellation.Cancel();
                throw;
            }
        }
    }
}
