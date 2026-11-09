using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Connects existing fixed-surface renderers to shared ordered selection, execution and persistence.</summary>
internal static class OfficeImageExportAdapter {
    internal static OfficeImageExportBatchOperation<TOptions> CreateOperation<TSurface, TOptions>(
        IReadOnlyList<TSurface> surfaces, TOptions options,
        Func<TSurface, int, TOptions, CancellationToken, OfficeImageExportResult> render)
        where TOptions : OfficeImageExportOptions {
        void RenderEach(TOptions effective, OfficeImageExportConsumer consumer, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            IReadOnlyList<int> selected = effective.PageSelector?.Resolve(surfaces.Count, effective.MaximumOutputCount)
                ?? Enumerable.Range(1, surfaces.Count).ToArray();
            if (selected.Count > effective.MaximumOutputCount) throw new InvalidOperationException("Surface selection exceeds the configured output limit.");
            for (int index = 0; index < selected.Count; index++) {
                token.ThrowIfCancellationRequested();
                int number = selected[index];
                consumer(render(surfaces[number - 1], number, effective, token).WithSequence(index, selected.Count));
            }
        }
        return new OfficeImageExportBatchOperation<TOptions>(options, options.Format, options.FileConflictPolicy,
            (_, effective) => {
                var results = new List<OfficeImageExportResult>();
                RenderEach(effective, results.Add, default);
                return results.AsReadOnly();
            }, exportEach: (_, effective, consumer, token) => RenderEach(effective, consumer, token));
    }
}
