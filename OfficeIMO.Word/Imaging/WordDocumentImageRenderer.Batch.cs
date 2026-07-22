using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        internal static IReadOnlyList<OfficeImageExportResult> RenderPages(WordDocument document,
            OfficeImageExportFormat format, WordImageExportOptions options) {
            var results = new List<OfficeImageExportResult>();
            RenderPages(document, format, options, results.Add);
            return results.AsReadOnly();
        }

        internal static void RenderPages(
            WordDocument document,
            OfficeImageExportFormat format,
            WordImageExportOptions options,
            OfficeImageExportConsumer consumer,
            CancellationToken cancellationToken = default) {
            if (consumer == null) throw new ArgumentNullException(nameof(consumer));
            IReadOnlyList<int> sectionPageCounts = EstimateSectionPageCounts(
                document,
                cancellationToken,
                options.CancellationCheckpoint);
            (int firstPage, int count) = ResolveBatchPageRange(options, sectionPageCounts);
            int[] pages = Enumerable.Range(firstPage, count).ToArray();
            OfficeImageExportBatchProcessor.ForEachOrdered(
                pages,
                options.MaximumDegreeOfParallelism,
                (pageIndex, _, token) => {
                    WordImageExportOptions pageOptions = options.Clone();
                    pageOptions.PageIndex = pageIndex;
                    WordDocumentVisualSnapshot snapshot = CreateSnapshot(
                        document,
                        pageOptions,
                        sectionPageCounts,
                        token);
                    token.ThrowIfCancellationRequested();
                    return RenderSnapshot(snapshot, format, pageOptions, token);
                },
                consumer,
                cancellationToken,
                options);
        }

        internal static IReadOnlyList<WordDocumentVisualSnapshot> CreateSnapshots(WordDocument document,
            WordImageExportOptions options,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<int> sectionPageCounts = EstimateSectionPageCounts(
                document,
                cancellationToken,
                options.CancellationCheckpoint);
            (int firstPage, int count) = ResolveBatchPageRange(options, sectionPageCounts);
            if (count > options.MaximumOutputCount) {
                throw new InvalidOperationException(
                    "Word page snapshot count exceeds " + nameof(OfficeImageExportOptions.MaximumOutputCount) +
                    ": " + count + " requested, " + options.MaximumOutputCount + " allowed.");
            }
            var snapshots = new List<WordDocumentVisualSnapshot>(count);
            for (int index = 0; index < count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                WordImageExportOptions pageOptions = options.Clone();
                pageOptions.PageIndex = firstPage + index;
                snapshots.Add(CreateSnapshot(
                    document,
                    pageOptions,
                    sectionPageCounts,
                    cancellationToken));
            }

            return snapshots.AsReadOnly();
        }

        private static (int FirstPage, int Count) ResolveBatchPageRange(WordImageExportOptions options,
            IReadOnlyList<int> sectionPageCounts) {
            int totalPages = Math.Max(1, sectionPageCounts.Sum());
            if (options.PageIndex >= totalPages) {
                throw new ArgumentOutOfRangeException(nameof(options),
                    $"First page index {options.PageIndex} is outside the estimated {totalPages}-page document.");
            }

            int available = totalPages - options.PageIndex;
            int count = options.PageCount.HasValue ? Math.Min(options.PageCount.Value, available) : available;
            return (options.PageIndex, count);
        }
    }
}
