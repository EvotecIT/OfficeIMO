using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        internal static IReadOnlyList<OfficeImageExportResult> RenderPages(WordDocument document,
            OfficeImageExportFormat format, WordImageExportOptions options) {
            IReadOnlyList<int> sectionPageCounts = EstimateSectionPageCounts(document);
            (int firstPage, int count) = ResolveBatchPageRange(options, sectionPageCounts);
            var results = new List<OfficeImageExportResult>(count);
            for (int index = 0; index < count; index++) {
                WordImageExportOptions pageOptions = options.Clone();
                pageOptions.PageIndex = firstPage + index;
                WordDocumentVisualSnapshot snapshot = CreateSnapshot(document, pageOptions, sectionPageCounts);
                results.Add(RenderSnapshot(snapshot, format, pageOptions));
            }

            return results.AsReadOnly();
        }

        internal static IReadOnlyList<WordDocumentVisualSnapshot> CreateSnapshots(WordDocument document,
            WordImageExportOptions options) {
            IReadOnlyList<int> sectionPageCounts = EstimateSectionPageCounts(document);
            (int firstPage, int count) = ResolveBatchPageRange(options, sectionPageCounts);
            var snapshots = new List<WordDocumentVisualSnapshot>(count);
            for (int index = 0; index < count; index++) {
                WordImageExportOptions pageOptions = options.Clone();
                pageOptions.PageIndex = firstPage + index;
                snapshots.Add(CreateSnapshot(document, pageOptions, sectionPageCounts));
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
