using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing;

/// <summary>Bounded, ordered batch processing shared by document image adapters.</summary>
public static class OfficeImageExportBatchProcessor {
    /// <summary>
    /// Executes a streaming batch under one render deadline and one aggregate budget tracker.
    /// When <paramref name="expectedOutputCount"/> is known, the count budget is checked before
    /// the producer starts so no partial output is emitted for a predictably oversized batch.
    /// </summary>
    public static void Run(
        OfficeImageExportOptions options,
        Action<OfficeImageExportConsumer, CancellationToken> producer,
        OfficeImageExportConsumer consumer,
        CancellationToken cancellationToken = default,
        int? expectedOutputCount = null) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (producer == null) throw new ArgumentNullException(nameof(producer));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        options.ValidateImageExportOptions();
        ValidateExpectedOutputCount(options, expectedOutputCount);
        cancellationToken.ThrowIfCancellationRequested();

        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            options.RenderTimeout,
            cancellationToken);
        try {
            RunWithinExecutionScope(options, producer, consumer, execution, expectedOutputCount);
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    /// <summary>
    /// Executes cancellable count discovery and a streaming batch under the same render deadline.
    /// This internal boundary lets provider adapters preflight predictable result counts without
    /// doing potentially expensive source traversal before the operation scope starts.
    /// </summary>
    internal static void RunWithPreflight(
        OfficeImageExportOptions options,
        Func<CancellationToken, int?> resolveExpectedOutputCount,
        Action<OfficeImageExportConsumer, CancellationToken> producer,
        OfficeImageExportConsumer consumer,
        CancellationToken cancellationToken = default) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (resolveExpectedOutputCount == null) throw new ArgumentNullException(nameof(resolveExpectedOutputCount));
        if (producer == null) throw new ArgumentNullException(nameof(producer));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        options.ValidateImageExportOptions();
        cancellationToken.ThrowIfCancellationRequested();

        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            options.RenderTimeout,
            cancellationToken);
        try {
            execution.ThrowIfCancellationRequested();
            int? expectedOutputCount = resolveExpectedOutputCount(execution.Token);
            execution.ThrowIfCancellationRequested();
            ValidateExpectedOutputCount(options, expectedOutputCount);
            RunWithinExecutionScope(options, producer, consumer, execution, expectedOutputCount);
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    /// <summary>
    /// Executes an asynchronous streaming batch under one render deadline and one aggregate budget tracker.
    /// When <paramref name="expectedOutputCount"/> is known, the count budget is checked before
    /// the producer starts so no partial output is emitted for a predictably oversized batch.
    /// </summary>
    public static async Task RunAsync(
        OfficeImageExportOptions options,
        Func<OfficeImageExportAsyncConsumer, CancellationToken, Task> producer,
        OfficeImageExportAsyncConsumer consumer,
        CancellationToken cancellationToken = default,
        int? expectedOutputCount = null) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (producer == null) throw new ArgumentNullException(nameof(producer));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        options.ValidateImageExportOptions();
        ValidateExpectedOutputCount(options, expectedOutputCount);
        cancellationToken.ThrowIfCancellationRequested();

        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            options.RenderTimeout,
            cancellationToken);
        try {
            OfficeImageExportAsyncConsumer accept = CreateGuardedAsyncConsumerCore(
                options,
                consumer,
                execution.Token,
                expectedOutputCount);
            await producer(accept, execution.Token).ConfigureAwait(false);
            execution.ThrowIfCancellationRequested();
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    /// <summary>
    /// Executes asynchronous count discovery and streaming under the same render deadline.
    /// Provider adapters use this when rendering must finish before the predictable result count is known.
    /// </summary>
    internal static async Task RunAsyncWithPreflight(
        OfficeImageExportOptions options,
        Func<CancellationToken, Task<int?>> resolveExpectedOutputCount,
        Func<OfficeImageExportAsyncConsumer, CancellationToken, Task> producer,
        OfficeImageExportAsyncConsumer consumer,
        CancellationToken cancellationToken = default) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (resolveExpectedOutputCount == null) throw new ArgumentNullException(nameof(resolveExpectedOutputCount));
        if (producer == null) throw new ArgumentNullException(nameof(producer));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        options.ValidateImageExportOptions();
        cancellationToken.ThrowIfCancellationRequested();

        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            options.RenderTimeout,
            cancellationToken);
        try {
            execution.ThrowIfCancellationRequested();
            int? expectedOutputCount = await resolveExpectedOutputCount(execution.Token).ConfigureAwait(false);
            execution.ThrowIfCancellationRequested();
            ValidateExpectedOutputCount(options, expectedOutputCount);
            OfficeImageExportAsyncConsumer accept = CreateGuardedAsyncConsumerCore(
                options,
                consumer,
                execution.Token,
                expectedOutputCount);
            await producer(accept, execution.Token).ConfigureAwait(false);
            execution.ThrowIfCancellationRequested();
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    /// <summary>
    /// Wraps a consumer with cancellation, diagnostic policy, and aggregate batch-budget enforcement.
    /// </summary>
    public static OfficeImageExportConsumer CreateGuardedConsumer(
        OfficeImageExportOptions options,
        OfficeImageExportConsumer consumer,
        CancellationToken cancellationToken = default) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        options.ValidateImageExportOptions();
        return CreateGuardedConsumerCore(options, consumer, cancellationToken, expectedOutputCount: null);
    }

    /// <summary>
    /// Wraps an asynchronous consumer with cancellation, diagnostic policy, and aggregate batch-budget enforcement.
    /// </summary>
    public static OfficeImageExportAsyncConsumer CreateGuardedAsyncConsumer(
        OfficeImageExportOptions options,
        OfficeImageExportAsyncConsumer consumer,
        CancellationToken cancellationToken = default) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        options.ValidateImageExportOptions();
        return CreateGuardedAsyncConsumerCore(options, consumer, cancellationToken, expectedOutputCount: null);
    }

    /// <summary>
    /// Renders items in bounded parallel windows and emits results in source order.
    /// A degree of one is strictly sequential.
    /// </summary>
    public static void ForEachOrdered<T>(
        IReadOnlyList<T> items,
        int maximumDegreeOfParallelism,
        Func<T, int, CancellationToken, OfficeImageExportResult> render,
        OfficeImageExportConsumer consumer,
        CancellationToken cancellationToken = default,
        OfficeImageExportOptions? options = null) {
        if (items == null) throw new ArgumentNullException(nameof(items));
        if (render == null) throw new ArgumentNullException(nameof(render));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        if (maximumDegreeOfParallelism < 1) throw new ArgumentOutOfRangeException(nameof(maximumDegreeOfParallelism));
        if (options != null) {
            Run(
                options,
                (accept, token) => ForEachOrderedCore(items, maximumDegreeOfParallelism, render, accept, token),
                consumer,
                cancellationToken,
                items.Count);
            return;
        }

        cancellationToken.ThrowIfCancellationRequested();
        ForEachOrderedCore(items, maximumDegreeOfParallelism, render, consumer, cancellationToken);
    }

    private static void ForEachOrderedCore<T>(
        IReadOnlyList<T> items,
        int maximumDegreeOfParallelism,
        Func<T, int, CancellationToken, OfficeImageExportResult> render,
        OfficeImageExportConsumer consumer,
        CancellationToken cancellationToken) {
        if (maximumDegreeOfParallelism == 1 || items.Count <= 1) {
            for (int index = 0; index < items.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                consumer(render(items[index], index, cancellationToken));
            }
            return;
        }

        for (int offset = 0; offset < items.Count; offset += maximumDegreeOfParallelism) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(maximumDegreeOfParallelism, items.Count - offset);
            var tasks = new Task<OfficeImageExportResult>[count];
            for (int localIndex = 0; localIndex < count; localIndex++) {
                int resolvedIndex = offset + localIndex;
                T item = items[resolvedIndex];
                tasks[localIndex] = Task.Run(
                    () => render(item, resolvedIndex, cancellationToken),
                    cancellationToken);
            }

            OfficeImageExportResult[] results = Task.WhenAll(tasks).GetAwaiter().GetResult();
            for (int localIndex = 0; localIndex < results.Length; localIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                consumer(results[localIndex]);
            }
        }
    }

    private static void ValidateExpectedOutputCount(
        OfficeImageExportOptions options,
        int? expectedOutputCount) {
        if (!expectedOutputCount.HasValue) return;
        if (expectedOutputCount.Value < 0) {
            throw new ArgumentOutOfRangeException(nameof(expectedOutputCount));
        }
        if (expectedOutputCount.Value > options.MaximumOutputCount) {
            throw new OfficeImageExportBatchLimitException(
                nameof(OfficeImageExportOptions.MaximumOutputCount),
                expectedOutputCount.Value,
                options.MaximumOutputCount);
        }
    }

    private static void RunWithinExecutionScope(
        OfficeImageExportOptions options,
        Action<OfficeImageExportConsumer, CancellationToken> producer,
        OfficeImageExportConsumer consumer,
        OfficeImageExportExecutionScope execution,
        int? expectedOutputCount) {
        OfficeImageExportConsumer accept = CreateGuardedConsumerCore(
            options,
            consumer,
            execution.Token,
            expectedOutputCount);
        producer(accept, execution.Token);
        execution.ThrowIfCancellationRequested();
    }

    private static OfficeImageExportConsumer CreateGuardedConsumerCore(
        OfficeImageExportOptions options,
        OfficeImageExportConsumer consumer,
        CancellationToken cancellationToken,
        int? expectedOutputCount) {
        var tracker = new OfficeImageExportBatchTracker(options);
        var gate = new object();
        return result => {
            lock (gate) {
                cancellationToken.ThrowIfCancellationRequested();
                if (result == null) throw new ArgumentNullException(nameof(result));
                result.Require(options.Policy);
                int sequenceIndex = tracker.Count;
                tracker.Add(result);
                consumer(result.WithSequence(sequenceIndex, expectedOutputCount));
            }
        };
    }

    private static OfficeImageExportAsyncConsumer CreateGuardedAsyncConsumerCore(
        OfficeImageExportOptions options,
        OfficeImageExportAsyncConsumer consumer,
        CancellationToken cancellationToken,
        int? expectedOutputCount) {
        var tracker = new OfficeImageExportBatchTracker(options);
        var gate = new SemaphoreSlim(1, 1);
        return async (result, token) => {
            using CancellationTokenSource? linked = CreateLinkedCancellationSource(cancellationToken, token);
            CancellationToken effectiveToken = linked?.Token ?? (token.CanBeCanceled ? token : cancellationToken);
            OfficeImageExportResult sequenced;
            await gate.WaitAsync(effectiveToken).ConfigureAwait(false);
            try {
                cancellationToken.ThrowIfCancellationRequested();
                token.ThrowIfCancellationRequested();
                if (result == null) throw new ArgumentNullException(nameof(result));
                result.Require(options.Policy);
                int sequenceIndex = tracker.Count;
                tracker.Add(result);
                sequenced = result.WithSequence(sequenceIndex, expectedOutputCount);
            } finally {
                gate.Release();
            }

            // Admission and sequence assignment are serialized, but arbitrary consumer code
            // must run outside the gate so a consumer can safely submit another result.
            await consumer(sequenced, token).ConfigureAwait(false);
        };
    }

    private static CancellationTokenSource? CreateLinkedCancellationSource(
        CancellationToken first,
        CancellationToken second) {
        if (!first.CanBeCanceled || !second.CanBeCanceled || first == second) return null;
        return CancellationTokenSource.CreateLinkedTokenSource(first, second);
    }
}
