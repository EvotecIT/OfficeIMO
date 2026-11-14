using OfficeIMO.TextAnalysis;
using Xunit;

namespace OfficeIMO.Tests;

public class TextOfficeTextIndexTests {
    // N = 3, average length 3, k1 = 1.2, b = 0.75. Expected scores were computed by hand from
    // idf = ln(1 + (N - df + 0.5) / (df + 0.5)) and idf · tf' · (k1 + 1) / (tf' + k1), tf' = tf / (1 - b + b · dl / avgdl).
    private static OfficeTextIndex TinyCorpus() {
        var index = new OfficeTextIndex();
        index.Add("d0", "apple banana");
        index.Add("d1", "apple apple cherry");
        index.Add("d2", "banana cherry cherry durian");
        return index;
    }

    [Fact]
    public void TermScoresMatchHandComputedBm25() {
        OfficeTextIndexSearchResult result = TinyCorpus().Search("apple");
        Assert.Equal(new[] { "d1", "d0" }, result.Hits.Select(static hit => hit.DocumentId));
        Assert.Equal(0.6462549902128865, result.Hits[0].Score, 12); // ln 1.6 · 2 · 2.2 / 3.2
        Assert.Equal(0.5442147286003256, result.Hits[1].Score, 12); // ln 1.6 · (4/3) · 2.2 / (4/3 + 1.2)
    }

    [Fact]
    public void PhraseUsesPhraseFrequencyAndSummedIdf() {
        OfficeTextIndexSearchResult result = TinyCorpus().Search("\"banana cherry\"");
        OfficeTextIndexHit hit = Assert.Single(result.Hits);
        Assert.Equal("d2", hit.DocumentId);
        Assert.Equal(0.8272063874724949, hit.Score, 12); // 2 · ln 1.6, tf 1, dl 4
        Assert.Equal(new[] { (0, 13) }, hit.Highlights.Select(static h => (h.Start, h.Length)));
        Assert.Empty(TinyCorpus().Search("\"cherry banana\"").Hits);
    }

    [Fact]
    public void PrefixExpandsToIndexTerms() {
        OfficeTextIndexSearchResult result = TinyCorpus().Search("CHER*");
        Assert.Equal(new[] { "d2", "d1" }, result.Hits.Select(static hit => hit.DocumentId));
        Assert.Equal(0.5908617053374964, result.Hits[0].Score, 12);
        Assert.Equal(0.47000362924573563, result.Hits[1].Score, 12);
        Assert.Equal(new[] { (7, 6), (14, 6) }, result.Hits[0].Highlights.Select(static h => (h.Start, h.Length)));
    }

    [Fact]
    public void BooleanOperatorsCombineClauses() {
        OfficeTextIndex index = TinyCorpus();
        Assert.Equal(new[] { "d1" }, index.Search("apple cherry").Hits.Select(static hit => hit.DocumentId));
        Assert.Equal(new[] { "d2", "d1", "d0" }, index.Search("apple OR durian").Hits.Select(static hit => hit.DocumentId));
        // d2: cherry (tf 2) plus the rare durian outweighs d1: cherry plus apple (tf 2).
        Assert.Equal(new[] { "d2", "d1" },
            index.Search("cherry AND (apple OR durian)").Hits.Select(static hit => hit.DocumentId));
        Assert.Equal(new[] { "d0", "d1", "d2" },
            index.Search("apple banana", new OfficeTextIndexSearchOptions { DefaultOperator = OfficeTextQueryOperator.Or })
                .Hits.Select(static hit => hit.DocumentId));
    }

    [Fact]
    public void EqualScoresKeepInsertionOrder() {
        var index = new OfficeTextIndex();
        for (int i = 0; i < 5; i++) index.Add("doc" + (4 - i), "same words here");
        Assert.Equal(new[] { "doc4", "doc3", "doc2", "doc1", "doc0" },
            index.Search("words").Hits.Select(static hit => hit.DocumentId));
    }

    [Fact]
    public void FieldWeightsRankTitleMatchesHigher() {
        var options = new OfficeTextIndexOptions();
        options.FieldWeights["title"] = 3;
        var index = new OfficeTextIndex(options);
        index.Add("body-match", new OfficeTextField("title", "Overview"), new OfficeTextField("body", "budget summary for the year"));
        index.Add("title-match", new OfficeTextField("title", "Budget"), new OfficeTextField("body", "summary for the year"));
        OfficeTextIndexSearchResult result = index.Search("budget");
        Assert.Equal(new[] { "title-match", "body-match" }, result.Hits.Select(static hit => hit.DocumentId));
        Assert.Equal("title", result.Hits[0].Highlights.Single().Field);
    }

    [Fact]
    public void AnalysisFoldsCaseNormalizationAndOptionallyAccents() {
        var analyzer = new OfficeTextAnalyzer();
        Assert.Equal(new[] { "strasse", "strasse", "finance", "σοφοσ", "σοφοσ" },
            analyzer.Analyze("STRASSE Straße ﬁnance ΣΟΦΟΣ σοφος").Select(static token => token.Term));
        Assert.Equal("ﬁnance", new OfficeTextAnalyzer(new OfficeTextAnalyzerOptions {
            Normalization = OfficeTextNormalization.None, CaseFolding = false
        }).NormalizeTerm("ﬁnance"));
        Assert.Equal("zażółć", analyzer.NormalizeTerm("ZAŻÓŁĆ"));
        var folding = new OfficeTextAnalyzer(new OfficeTextAnalyzerOptions { AccentFolding = true });
        Assert.Equal("zazolc", folding.NormalizeTerm("Zażółć"));
        Assert.Equal("resume", folding.NormalizeTerm("Résumé"));
    }

    [Fact]
    public void AnalyzerReportsSourceOffsetsAndPositions() {
        IReadOnlyList<OfficeTextToken> tokens = new OfficeTextAnalyzer().Analyze("  Hello,  wörld e-mail");
        Assert.Equal(new[] { ("hello", 2, 5, 0), ("wörld", 10, 5, 1), ("e", 16, 1, 2), ("mail", 18, 4, 3) },
            tokens.Select(static token => (token.Term, token.Start, token.Length, token.Position)));
    }

    [Fact]
    public void MultiTokenWordsMatchAsPhrases() {
        var index = new OfficeTextIndex();
        index.Add("hyphen", "send an e-mail today");
        index.Add("split", "mail the e form");
        Assert.Equal(new[] { "hyphen" }, index.Search("e-mail").Hits.Select(static hit => hit.DocumentId));
        Assert.Equal(new[] { "hyphen", "split" }, index.Search("E MAIL").Hits.Select(static hit => hit.DocumentId).OrderBy(static id => id));
    }

    [Fact]
    public void PunctuationOnlyQueriesHaveNoSearchableText() {
        OfficeTextIndexSearchResult result = TinyCorpus().Search("!!! --");
        Assert.False(result.HasSearchableText);
        Assert.Empty(result.Hits);
        Assert.True(TinyCorpus().Search("apple").HasSearchableText);
    }

    [Fact]
    public void LimitsAreReportedNotExceeded() {
        var index = new OfficeTextIndex(new OfficeTextIndexOptions { MaximumTokens = 5, MaximumDocuments = 10 });
        Assert.True(index.Add("a", "one two three"));
        Assert.True(index.Add("b", "four five six"));
        Assert.False(index.Add("c", "seven"));
        Assert.True(index.LimitReached);
        Assert.Equal(5, index.TokenCount);
        Assert.Empty(index.Search("six").Hits);

        var terms = new OfficeTextIndex();
        terms.Add("t", "alpha alpine altitude alto");
        OfficeTextIndexSearchResult limited = terms.Search("al*", new OfficeTextIndexSearchOptions { MaximumPrefixExpansions = 2 });
        Assert.True(limited.PrefixExpansionLimited);
        Assert.Throws<ArgumentException>(() => terms.Search("a b c", new OfficeTextIndexSearchOptions { MaximumClauses = 2 }));
        OfficeTextIndexSearchResult work = terms.Search("alpha", new OfficeTextIndexSearchOptions { MaximumPostingsVisited = 1 });
        Assert.False(work.WorkLimitReached);
        Assert.True(TinyCorpus().Search("apple", new OfficeTextIndexSearchOptions { MaximumPostingsVisited = 1 }).WorkLimitReached);
    }

    [Fact]
    public void SnippetCentresOnHighlightWithRelativeOffsets() {
        var index = new OfficeTextIndex();
        index.Add("long", "Lorem ipsum dolor sit amet, the quarterly revenue grew by twelve percent in the north region.");
        OfficeTextIndexHit hit = index.Search("revenue").Hits.Single();
        OfficeTextSnippet snippet = index.CreateSnippet(hit, "body", 30)!;
        Assert.True(snippet.StartsMidText);
        Assert.True(snippet.EndsMidText);
        OfficeTextHighlight highlight = Assert.Single(snippet.Highlights);
        Assert.Equal("revenue", snippet.Text.Substring(highlight.Start, highlight.Length));
        Assert.Null(index.CreateSnippet(hit, "missing"));
    }

    [Fact]
    public void QueryParserBuildsTheExpectedTree() {
        Assert.Equal("(a AND \"b c\" AND d*)", OfficeTextQuery.Parse("a \"b c\" d*")!.ToString());
        Assert.Equal("((a AND b) OR c)", OfficeTextQuery.Parse("a b OR c")!.ToString());
        Assert.Equal("(a AND (b OR c))", OfficeTextQuery.Parse("a (b OR c)")!.ToString());
        Assert.Equal("((a AND b) OR c)", OfficeTextQuery.Parse("a AND b c", OfficeTextQueryOperator.Or)!.ToString());
        Assert.Null(OfficeTextQuery.Parse("  AND OR \"\" "));
    }
}
