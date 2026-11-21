namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    /// <summary>
    /// True when extraction must use the catalog-preserving page organizer: the subset catalog writer rebuilds only the
    /// catalog entries preflight proved it can copy, and every outline tree must keep its items individually.
    /// </summary>
    internal static bool UsesCatalogPreservingWriter(PdfReadDocument document, IReadOnlyList<PdfRewriteBlocker> rewriteBlockers) =>
        PdfTaggedPageOrganizer.IsTagged(document) || document.FormFields.Count > 0 || HasCatalogOutlines(document) ||
        rewriteBlockers.Any(static blocker => blocker.Kind != PdfRewriteBlockerKind.Encryption);

    /// <summary>
    /// Walks everything the catalog reaches, as the catalog-preserving writer does before it writes, and returns why a
    /// reference to a missing object or to another object generation makes that graph unsafe to rewrite.
    /// </summary>
    internal static string? FindCatalogObjectGraphFailure(PdfReadDocument document, CancellationToken cancellationToken) {
        if (document.Security.RootObjectNumber is not int root || !document.Objects.TryGetValue(root, out var catalog))
            return "PDF object graph is not safe for rewriting by OfficeIMO.Pdf yet: the document catalog was not found.";
        try {
            new ObjectCollector(document.Objects, cancellationToken: cancellationToken)
                .CollectObjectGraph(new PdfReference(catalog.ObjectNumber, catalog.Generation));
            return null;
        } catch (InvalidOperationException ex) {
            return "PDF object graph is not safe for rewriting by OfficeIMO.Pdf yet: " + ex.Message;
        }
    }

    private static bool HasCatalogOutlines(PdfReadDocument document) =>
        document.Security.RootObjectNumber is int root && document.Objects.TryGetValue(root, out var catalog) &&
        catalog.Value is PdfDictionary dictionary && dictionary.Items.ContainsKey("Outlines");

    /// <summary>One validated input and catalog snapshot shared by a compound page extraction.</summary>
    internal sealed class ExtractionSession {
        private readonly PdfReadDocument _document;
        private readonly PdfMutationPlan _plan;
        private PdfDocumentProperties? _metadata;
        private readonly CatalogRewriteState _catalog;
        private readonly PdfFileVersion _fileVersion;
        private readonly byte[] _source;
        private readonly PdfLoadOptions? _options;
        private readonly bool _preservesCatalog;

        private readonly CancellationToken _cancellationToken;

        internal ExtractionSession(
            byte[] pdf,
            PdfLoadOptions? options,
            Func<PdfReadDocument>? documentFactory = null,
            CancellationToken cancellationToken = default) {
            Guard.NotNull(pdf, nameof(pdf));
            cancellationToken.ThrowIfCancellationRequested();
            _cancellationToken = cancellationToken;
            _source = pdf;
            _options = options;
            (_plan, _document) = documentFactory == null
                ? PdfMutationPlanner.RequireFullRewriteDocument(pdf, PdfMutationOperation.ExtractPages, options, cancellationToken: cancellationToken)
                : PdfMutationPlanner.RequireFullRewriteDocument(pdf, PdfMutationOperation.ExtractPages, documentFactory, options, cancellationToken: cancellationToken);
            _catalog = _document.GetCatalogRewriteState(cancellationToken);
            _fileVersion = GetSourceFileVersion(pdf);
            // The subset catalog writer rebuilds only the catalog entries preflight proved it can copy. Anything else,
            // and every outline tree (whose items must survive individually), stays with the catalog-preserving writer.
            _preservesCatalog = UsesCatalogPreservingWriter(_document, _plan.Preflight.RewriteBlockers);
        }

        internal int PageCount => _document.Pages.Count;

        /// <summary>Extracts pages, adding a sentence to <paramref name="report"/> for each navigation item it removes or changes.</summary>
        internal byte[] Extract(int[] pageNumbers, long? maximumOutputBytes = null, ICollection<string>? report = null) {
            _cancellationToken.ThrowIfCancellationRequested();
            Guard.NotNull(pageNumbers, nameof(pageNumbers));
            if (pageNumbers.Length == 0) throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
            ValidatePageNumbers(pageNumbers, PageCount, nameof(pageNumbers), _cancellationToken);
            if (_preservesCatalog) {
                try {
                    return PdfTaggedPageOrganizer.Rewrite(_source, _document, pageNumbers, _options, maximumOutputBytes, token: _cancellationToken, report: report);
                } catch (InvalidOperationException exception) {
                    // The organizer refuses structure, navigation and form graphs it cannot retarget without losing
                    // or misattributing content; report that as a blocked mutation rather than an internal failure.
                    throw PdfMutationBlockedException.PolicyConflict(_plan, "FullRewrite.TaggedPageOrganization", exception.Message);
                }
            }
            var objects = new int[pageNumbers.Length];
            for (int index = 0; index < pageNumbers.Length; index++) {
                _cancellationToken.ThrowIfCancellationRequested();
                objects[index] = _document.Pages[pageNumbers[index] - 1].ObjectNumber;
            }
            try {
                return ExtractPages(_document.Objects, _metadata ??= _document.GetUncheckedMetadata(_cancellationToken), objects,
                    catalogState: _catalog, fileVersion: _fileVersion, maximumOutputBytes: maximumOutputBytes,
                    cancellationToken: _cancellationToken, report: report);
            } catch (InvalidOperationException exception) {
                // A selected page can reach a missing or mismatched object that the page-graph preflight did not visit;
                // refuse it the same way preflight refuses invalid references.
                throw PdfMutationBlockedException.PolicyConflict(_plan,
                    "FullRewrite." + PdfRewriteBlockerKind.InvalidObjectReferences, exception.Message);
            }
        }

        internal byte[] Extract(PdfPageRange[] ranges, ICollection<string>? report = null) {
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
            return Extract(pages, report: report);
        }
    }
}
