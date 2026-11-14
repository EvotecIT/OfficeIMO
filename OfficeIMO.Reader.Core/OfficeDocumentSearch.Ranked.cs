using OfficeIMO.Reader.Internal.TextAnalysis;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentReadResultExtensions {
    // Returns null when the query has no searchable words or nothing matches, so the caller runs literal search.
    private static OfficeDocumentSearchResult? SearchRanked(
        OfficeDocumentReadResult document,
        string query,
        OfficeDocumentSearchOptions options,
        CancellationToken cancellationToken) {
        OfficeDocumentBlock[] blocks = (document.Blocks ?? Array.Empty<OfficeDocumentBlock>()).ToArray();
        OfficeTextIndex index = ReaderRankedSearch.CreateIndex(options.Ranking, options.MatchCase);
        for (int ordinal = 0; ordinal < blocks.Length; ordinal++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentBlock block = blocks[ordinal];
            // A heading's own path repeats its text; indexing it twice would rank bare headings above their content.
            string? headingPath = string.Equals(block.Kind, "heading", StringComparison.OrdinalIgnoreCase)
                ? null
                : block.Location?.HeadingPath;
            if (!ReaderRankedSearch.Add(index, ordinal, headingPath, block.Text)) break;
        }

        cancellationToken.ThrowIfCancellationRequested();
        OfficeTextIndexSearchResult ranked = ReaderRankedSearch.Search(index, query, options.Ranking, options.MaximumResults);
        if (!ranked.HasSearchableText || ranked.Hits.Count == 0) return null;

        var hits = new List<OfficeDocumentSearchHit>(ranked.Hits.Count);
        foreach (OfficeTextIndexHit rankedHit in ranked.Hits) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentBlock block = blocks[rankedHit.DocumentIndex];
            IReadOnlyList<OfficeDocumentSearchHighlight> highlights = ReaderRankedSearch.BodyHighlights(rankedHit);
            int start = highlights.Count > 0 ? highlights[0].Start : 0;
            int length = highlights.Count > 0 ? highlights[0].Length : 0;
            hits.Add(new OfficeDocumentSearchHit(
                block,
                start,
                length,
                LocateRankedHit(document, block, start, length),
                rankedHit.Score,
                highlights));
        }

        return new OfficeDocumentSearchResult(
            document.Source ?? new OfficeDocumentSource(),
            query,
            document.GetTotalPageCount(),
            hits.AsReadOnly(),
            maximumResultsReached: ranked.TotalMatches > hits.Count || ranked.WorkLimitReached,
            ranked: true);
    }

    // Pages holding the first matched range when the page fragments reproduce the block text exactly;
    // otherwise every page the block spans.
    private static IReadOnlyList<OfficeDocumentPageLocation> LocateRankedHit(
        OfficeDocumentReadResult document,
        OfficeDocumentBlock block,
        int start,
        int length) {
        IReadOnlyList<OfficeDocumentPageLocation> locations = document.Locate(block);
        string text = block.Text ?? string.Empty;
        if (locations.Count == 0 || length == 0) return locations;
        IReadOnlyList<IReadOnlyList<OfficeDocumentPageLocation>>? mapped =
            MapOccurrencesByFragmentOffsets(locations, block.Id, text, length, new[] { start }, 1);
        return mapped != null && mapped.Count == 1 ? mapped[0] : locations;
    }
}
