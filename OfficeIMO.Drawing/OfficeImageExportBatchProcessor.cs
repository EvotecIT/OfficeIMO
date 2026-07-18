using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing;

/// <summary>Bounded, ordered batch processing shared by document image adapters.</summary>
public static class OfficeImageExportBatchProcessor {
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
        var tracker = new OfficeImageExportBatchTracker(options);
        return result => {
            cancellationToken.ThrowIfCancellationRequested();
            if (result == null) throw new ArgumentNullException(nameof(result));
            result.Require(options.Policy);
            tracker.Add(result);
            consumer(result);
        };
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
        var tracker = new OfficeImageExportBatchTracker(options);
        return async (result, token) => {
            cancellationToken.ThrowIfCancellationRequested();
            token.ThrowIfCancellationRequested();
            if (result == null) throw new ArgumentNullException(nameof(result));
            result.Require(options.Policy);
            tracker.Add(result);
            await consumer(result, token).ConfigureAwait(false);
        };
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
        OfficeImageExportConsumer accept = options == null
            ? consumer
            : CreateGuardedConsumer(options, consumer, cancellationToken);

        if (maximumDegreeOfParallelism == 1 || items.Count <= 1) {
            for (int index = 0; index < items.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                accept(render(items[index], index, cancellationToken));
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
                accept(results[localIndex]);
            }
        }
    }
}
