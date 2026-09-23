using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    /// <summary>One validated input and catalog snapshot shared by a compound page extraction.</summary>
    internal sealed class ExtractionSession {
        private readonly PdfReadDocument _document;
        private PdfMetadata? _metadata;
        private readonly CatalogRewriteState _catalog;
        private readonly PdfFileVersion _fileVersion;

        private readonly CancellationToken _cancellationToken;

        internal ExtractionSession(
            byte[] pdf,
            PdfLoadOptions? options,
            Func<PdfReadDocument>? documentFactory = null,
            CancellationToken cancellationToken = default) {
            Guard.NotNull(pdf, nameof(pdf));
            cancellationToken.ThrowIfCancellationRequested();
            _cancellationToken = cancellationToken;
            (_, _document) = documentFactory == null
                ? PdfMutationPlanner.RequireFullRewriteDocument(pdf, PdfMutationOperation.ExtractPages, options, cancellationToken: cancellationToken)
                : PdfMutationPlanner.RequireFullRewriteDocument(pdf, PdfMutationOperation.ExtractPages, documentFactory, options, cancellationToken: cancellationToken);
            _catalog = ExtractCatalogRewriteState(_document.Objects, _document.TrailerRaw, cancellationToken);
            _fileVersion = GetSourceFileVersion(pdf);
        }

        internal int PageCount => _document.Pages.Count;

        internal byte[] Extract(int[] pageNumbers, long? maximumOutputBytes = null) {
            _cancellationToken.ThrowIfCancellationRequested();
            Guard.NotNull(pageNumbers, nameof(pageNumbers));
            if (pageNumbers.Length == 0) throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
            ValidatePageNumbers(pageNumbers, PageCount, nameof(pageNumbers), _cancellationToken);
            var objects = new int[pageNumbers.Length];
            for (int index = 0; index < pageNumbers.Length; index++) {
                _cancellationToken.ThrowIfCancellationRequested();
                objects[index] = _document.Pages[pageNumbers[index] - 1].ObjectNumber;
            }
            return ExtractPages(_document.Objects, _metadata ??= _document.UncheckedMetadata, objects,
                catalogState: _catalog, fileVersion: _fileVersion, maximumOutputBytes: maximumOutputBytes,
                cancellationToken: _cancellationToken);
        }

        internal byte[] Extract(PdfPageRange[] ranges) {
            _cancellationToken.ThrowIfCancellationRequested();
            Guard.NotNull(ranges, nameof(ranges));
            ValidatePageRanges(ranges, PageCount, "pageRanges", _cancellationToken);
            long totalPages = 0L;
            foreach (PdfPageRange range in ranges) {
                _cancellationToken.ThrowIfCancellationRequested();
                totalPages += range.PageCount;
                if (totalPages > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(ranges), "Selected page count exceeds the supported in-memory size.");
            }
            int[] pages = new int[(int)totalPages];
            int index = 0;
            foreach (PdfPageRange range in ranges) {
                _cancellationToken.ThrowIfCancellationRequested();
                for (int number = range.FirstPage; number <= range.LastPage; number++) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    pages[index++] = number;
                }
            }
            return Extract(pages);
        }
    }
}
