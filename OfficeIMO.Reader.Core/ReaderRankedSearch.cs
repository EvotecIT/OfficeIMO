using OfficeIMO.Reader.Internal.TextAnalysis;

namespace OfficeIMO.Reader;

/// <summary>Builds the shared ranked index over blocks or chunks: a heading field and a body field.</summary>
internal static class ReaderRankedSearch {
    internal const string HeadingField = "heading";
    internal const string BodyField = "body";

    internal static OfficeTextIndex CreateIndex(OfficeDocumentRankingOptions? ranking, bool matchCase) {
        OfficeDocumentRankingOptions effective = ranking ?? new OfficeDocumentRankingOptions();
        var options = new OfficeTextIndexOptions {
            K1 = effective.K1,
            B = effective.B,
            MaximumTokens = effective.MaximumIndexedTokens,
            Analyzer = new OfficeTextAnalyzerOptions {
                CaseFolding = !matchCase,
                AccentFolding = effective.AccentFolding
            }
        };
        options.FieldWeights[HeadingField] = effective.HeadingWeight;
        options.FieldWeights[BodyField] = effective.BodyWeight;
        return new OfficeTextIndex(options);
    }

    internal static bool Add(OfficeTextIndex index, int ordinal, string? heading, string? body) =>
        string.IsNullOrEmpty(heading)
            ? index.Add(ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture),
                new OfficeTextField(BodyField, body))
            : index.Add(ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture),
                new OfficeTextField(HeadingField, heading), new OfficeTextField(BodyField, body));

    internal static OfficeTextIndexSearchResult Search(OfficeTextIndex index, string query,
        OfficeDocumentRankingOptions? ranking, int maximumResults) {
        OfficeDocumentRankingOptions effective = ranking ?? new OfficeDocumentRankingOptions();
        OfficeTextQuery? parsed = OfficeTextQuery.Parse(query, effective.DefaultOperator == OfficeDocumentQueryOperator.Or
            ? OfficeTextQueryOperator.Or
            : OfficeTextQueryOperator.And);
        if (parsed == null) return index.Search(string.Empty);
        return index.Search(parsed, new OfficeTextIndexSearchOptions {
            MaximumResults = maximumResults,
            MaximumClauses = effective.MaximumClauses,
            MaximumPrefixExpansions = effective.MaximumPrefixExpansions,
            MaximumPostingsVisited = effective.MaximumPostingsVisited
        });
    }

    internal static IReadOnlyList<OfficeDocumentSearchHighlight> BodyHighlights(OfficeTextIndexHit hit) =>
        hit.Highlights
            .Where(static highlight => string.Equals(highlight.Field, BodyField, StringComparison.Ordinal))
            .Select(static highlight => new OfficeDocumentSearchHighlight(highlight.Start, highlight.Length))
            .ToArray();
}
