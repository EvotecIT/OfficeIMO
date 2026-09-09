namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    /// <summary>One validated input and catalog snapshot shared by a compound page extraction.</summary>
    internal sealed class ExtractionSession {
        private readonly PdfReadDocument _document;
        private readonly CatalogRewriteState _catalog;
        private readonly PdfFileVersion _fileVersion;

        internal ExtractionSession(byte[] pdf, PdfLoadOptions? options, Func<PdfReadDocument>? documentFactory = null) {
            Guard.NotNull(pdf, nameof(pdf));
            (_, _document) = documentFactory == null
                ? PdfMutationPlanner.RequireFullRewriteDocument(pdf, PdfMutationOperation.ExtractPages, options)
                : PdfMutationPlanner.RequireFullRewriteDocument(pdf, PdfMutationOperation.ExtractPages, documentFactory, options);
            _catalog = ExtractCatalogRewriteState(_document.Objects, _document.TrailerRaw);
            _fileVersion = GetSourceFileVersion(pdf);
        }

        internal int PageCount => _document.Pages.Count;

        internal byte[] Extract(int[] pageNumbers, long? maximumOutputBytes = null) {
            Guard.NotNull(pageNumbers, nameof(pageNumbers));
            if (pageNumbers.Length == 0) throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
            ValidatePageNumbers(pageNumbers, PageCount, nameof(pageNumbers));
            int[] objects = pageNumbers.Select(number => _document.Pages[number - 1].ObjectNumber).ToArray();
            return ExtractPages(_document.Objects, _document.UncheckedMetadata, objects,
                catalogState: _catalog, fileVersion: _fileVersion, maximumOutputBytes: maximumOutputBytes);
        }

        internal byte[] Extract(PdfPageRange[] ranges) {
            Guard.NotNull(ranges, nameof(ranges));
            ValidatePageRanges(ranges, PageCount, "pageRanges");
            int[] pages = new int[ranges.Sum(static range => range.PageCount)];
            int index = 0;
            foreach (PdfPageRange range in ranges) {
                for (int number = range.FirstPage; number <= range.LastPage; number++) pages[index++] = number;
            }
            return Extract(pages);
        }
    }
}
