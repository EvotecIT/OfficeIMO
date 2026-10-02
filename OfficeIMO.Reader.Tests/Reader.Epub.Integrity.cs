using OfficeIMO.Epub;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;
using Xunit;
using static OfficeIMO.Shared.Tests.EpubIntegrityFixtures;

namespace OfficeIMO.Tests;

public sealed class ReaderEpubIntegrityTests {
    [Theory]
    [InlineData("count", "epub.chapter.count-limit")]
    [InlineData("text", "epub.chapter.text-total-limit")]
    public void RichReader_ExposesChapterLossAndRespectsCallerBudgets(string budget, string code) {
        byte[] package = Package(new[] {
            ("a", "a.xhtml", "application/xhtml+xml", ""), ("b", "b.xhtml", "application/xhtml+xml", "") },
            "<itemref idref='a'/><itemref idref='b'/>", new[] { ("a.xhtml", Xhtml("<p>One</p>")), ("b.xhtml", Xhtml("<p>Two</p>")) });
        var options = new EpubReadOptions {
            MaxChapters = budget == "count" ? 1 : 500,
            MaxTotalTextCharacters = budget == "text" ? 3 : 1000
        };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddEpubHandler(options).Build();
        OfficeDocumentReadResult result = reader.ReadDocument(new MemoryStream(package), "book.epub");
        Assert.Single(result.Pages);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == code && diagnostic.Category == OfficeDocumentDiagnosticCategory.Limit);
        Assert.Equal("2", Assert.Single(result.Metadata, entry => entry.Name == "RequestedChapterCount").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, entry => entry.Name == "SkippedChapterCount").Value);
        Assert.Equal("False", Assert.Single(result.Metadata, entry => entry.Name == "ChaptersComplete").Value);
        Assert.False(options.IncludeRawHtml);
    }

    [Fact]
    public void RichReader_PreservesRepeatedChapterPagesAndUniqueBlockIds() {
        byte[] package = Package(new[] { ("c", "chapter.xhtml", "application/xhtml+xml", "") },
            "<itemref idref='c'/><itemref idref='c'/>", new[] { ("chapter.xhtml", Xhtml("<p>Repeated</p>")) });
        OfficeDocumentReadResult result = EpubReaderAdapter.ReadDocument(new MemoryStream(package), "book.epub");
        Assert.Equal(2, result.Pages.Count);
        Assert.All(result.Pages, page => Assert.NotEmpty(page.Blocks));
        Assert.Equal(result.Blocks.Count, result.Blocks.Select(block => block.Id).Distinct().Count());
        Assert.Equal("True", Assert.Single(result.Metadata, entry => entry.Name == "ChaptersComplete").Value);
    }
}
