using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderRankedSearchTests {
    private static readonly OfficeDocumentSearchOptions Ranked = new() { Mode = OfficeDocumentSearchMode.Ranked };

    private static OfficeDocumentReadResult WordReport() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph("Travel policy").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Employees book travel through the portal. Revenue targets do not change travel rules.");
            document.AddParagraph("Meals are reimbursed up to the daily limit.");
            document.AddParagraph("Quarterly revenue").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Revenue grew in every region. Subscription revenue and services revenue both rose.");
            document.AddParagraph("The north region reported record bookings.");
            document.Save();
        }
        stream.Position = 0;
        return OfficeIMO.Reader.Tests.ReaderTestReaders.Word().ReadDocument(stream, "report.docx");
    }

    [Fact]
    public void RankedWordSearchPutsTheDensestBlockFirst() {
        OfficeDocumentReadResult document = WordReport();
        OfficeDocumentSearchResult result = document.Search("revenue", Ranked);

        Assert.True(result.Ranked);
        Assert.StartsWith("Revenue grew in every region.", result.Hits[0].Block.Text, StringComparison.Ordinal);
        Assert.Equal(result.Hits.Select(static hit => hit.Score!.Value).OrderByDescending(static score => score),
            result.Hits.Select(static hit => hit.Score!.Value));
        OfficeDocumentSearchHit best = result.Hits[0];
        Assert.Equal(3, best.Highlights.Count);
        Assert.All(best.Highlights, highlight =>
            Assert.Equal("revenue", best.Block.Text.Substring(highlight.Start, highlight.Length), StringComparer.OrdinalIgnoreCase));
        Assert.Equal(best.Highlights[0].Start, best.StartIndex);
        Assert.Contains(result.Hits, hit => hit.Block.Text.StartsWith("The north region", StringComparison.Ordinal));
    }

    [Fact]
    public void HeadingPathMatchesRankBlocksUnderThatHeading() {
        OfficeDocumentSearchResult result = WordReport().Search("quarterly", Ranked);
        Assert.NotEmpty(result.Hits);
        Assert.All(result.Hits, hit => Assert.Contains("Quarterly", hit.Block.Location.HeadingPath ?? hit.Block.Text, StringComparison.Ordinal));
        Assert.Contains(result.Hits, hit => hit.Highlights.Count == 0); // matched through the heading path only
    }

    [Fact]
    public void PhrasePrefixAndOrQueriesWorkOnDocuments() {
        OfficeDocumentReadResult document = WordReport();
        OfficeDocumentSearchHit phrase = Assert.Single(document.Search("\"services revenue\"", Ranked).Hits);
        Assert.Equal("services revenue", phrase.Block.Text.Substring(phrase.StartIndex, phrase.Length));

        Assert.Contains(document.Search("reimburs*", Ranked).Hits, hit => hit.Block.Text.StartsWith("Meals", StringComparison.Ordinal));
        string[] either = document.Search("meals OR bookings", Ranked).Hits.Select(static hit => hit.Block.Text.Substring(0, 5)).ToArray();
        Assert.Equal(new[] { "Meals", "The n" }, either.OrderBy(static text => text, StringComparer.Ordinal));
        Assert.Empty(document.Search("meals bookings", Ranked).Hits.Where(static hit => hit.Score.HasValue));
    }

    [Fact]
    public void QueriesWithoutRankedMatchesFallBackToLiteralSearch() {
        OfficeDocumentReadResult document = WordReport();
        OfficeDocumentSearchResult substring = document.Search("venu", Ranked);
        Assert.False(substring.Ranked);
        Assert.NotEmpty(substring.Hits);
        Assert.All(substring.Hits, hit => Assert.Null(hit.Score));

        OfficeDocumentSearchResult punctuation = document.Search(".", Ranked);
        Assert.False(punctuation.Ranked);
        Assert.Equal(document.Search(".").Hits.Count, punctuation.Hits.Count);
    }

    [Fact]
    public void MatchCaseKeepsRankedTermsCaseSensitive() {
        OfficeDocumentReadResult document = WordReport();
        var options = new OfficeDocumentSearchOptions { Mode = OfficeDocumentSearchMode.Ranked, MatchCase = true };
        OfficeDocumentSearchResult result = document.Search("Revenue", options);
        Assert.True(result.Ranked);
        Assert.All(result.Hits.SelectMany(static hit => hit.Highlights.Select(h => hit.Block.Text.Substring(h.Start, h.Length))),
            text => Assert.Equal("Revenue", text));
    }

    [Fact]
    public void RankedPdfHitsKeepTheirPageCitations() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.AddText("Budget overview for the year."))
            .PageBreak()
            .Paragraph(paragraph => paragraph.AddText("The needle appears here, and a second needle follows."))
            .PageBreak()
            .Paragraph(paragraph => paragraph.AddText("Closing notes mention the needle once."))
            .ToBytes();
        OfficeDocumentReadResult document = PdfReaderAdapter.ReadDocument(pdf, sourceName: "pages.pdf");

        OfficeDocumentSearchResult result = document.Search("NEEDLE", Ranked);

        Assert.True(result.Ranked);
        Assert.Equal(2, result.Hits.Count);
        Assert.Equal(new[] { 2 }, result.Hits[0].Pages.Select(static page => page.Number!.Value));
        Assert.Equal(new[] { 3 }, result.Hits[1].Pages.Select(static page => page.Number!.Value));
        Assert.Equal(new[] { 2, 3 }, result.PageNumbers);
    }

    [Fact]
    public void MaximumResultsBoundsRankedHits() {
        OfficeDocumentSearchResult result = WordReport().Search("revenue OR travel", new OfficeDocumentSearchOptions {
            Mode = OfficeDocumentSearchMode.Ranked, MaximumResults = 1
        });
        Assert.Single(result.Hits);
        Assert.True(result.MaximumResultsReached);
    }

    [Fact]
    public void ChunkIndexRanksChunksAndKeepsTieOrder() {
        var chunks = new[] {
            Chunk("c0", "General", "Nothing relevant here."),
            Chunk("c1", "Pricing", "Prices rise in March. Pricing tables follow."),
            Chunk("c2", "Notes", "Same text."),
            Chunk("c3", "Notes", "Same text."),
            Chunk("c4", "Prices", "The price list is attached.")
        };
        var index = new ReaderChunkSearchIndex(chunks);

        Assert.Equal(5, index.Count);
        ReaderChunkSearchResult prices = index.Search("pric*");
        Assert.Equal(new[] { "c1", "c4" }, prices.Hits.Select(static hit => hit.Chunk.Id));
        Assert.Equal(new[] { "c2", "c3" }, index.Search("same text").Hits.Select(static hit => hit.Chunk.Id));
        Assert.False(index.Search("?!").HasSearchableText);

        ReaderChunkSearchHit heading = index.Search("notes").Hits[0];
        Assert.Empty(heading.Highlights);
    }

    [Fact]
    public void ChunkIndexWorksOnChunksFromAGeneratedDocument() {
        OfficeDocumentReadResult document = WordReport();
        var index = new ReaderChunkSearchIndex(document.Chunks);
        ReaderChunkSearchResult result = index.Search("subscription revenue");
        ReaderChunkSearchHit best = result.Hits[0];
        Assert.Contains("Subscription revenue", best.Chunk.Text, StringComparison.Ordinal);
        Assert.All(best.Highlights, highlight => Assert.InRange(highlight.Start + highlight.Length, 1, best.Chunk.Text.Length));
    }

    private static ReaderChunk Chunk(string id, string heading, string text) => new() {
        Id = id,
        Text = text,
        Location = new ReaderLocation { HeadingPath = heading }
    };
}
