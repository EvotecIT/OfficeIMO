using System.Text;
using OfficeIMO.Epub;
using Xunit;
using static OfficeIMO.Shared.Tests.EpubIntegrityFixtures;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubReaderFeedbackContracts {
    [Theory]
    [InlineData("malformed", "epub.chapter.invalid-xhtml")]
    [InlineData("oversized", "epub.chapter.size-limit")]
    [InlineData("missing", "epub.spine.resource-missing")]
    [InlineData("unsupported", "epub.spine.unsupported-media-type")]
    public void Load_RepeatedFailedPositionsKeepDiagnosticsBounded(string failure, string code) {
        const int positions = 10_000;
        byte[] package = Package(new[] { ("c", "chapter.xhtml", failure == "unsupported" ? "text/plain" : "application/xhtml+xml", "") },
            string.Concat(Enumerable.Repeat("<itemref idref='c'/>", positions)), failure == "missing" ? Array.Empty<(string, string)>() :
            new[] { ("chapter.xhtml", failure == "malformed" ? "<html>" : Xhtml("<p>Body</p>")) });
        EpubDocument book = EpubDocument.Load(new MemoryStream(package), new EpubReadOptions {
            MaxChapterBytes = failure == "oversized" ? 1 : null });
        Assert.Empty(book.Chapters);
        Assert.Equal(positions, book.ReadSummary.RequestedChapterCount);
        Assert.Equal(positions, book.ReadSummary.SkippedChapterCount);
        Assert.Single(book.Diagnostics, diagnostic => diagnostic.Code == code);
        Assert.InRange(book.Diagnostics.Count, 1, 10);
    }

    [Fact]
    public void Load_ArchiveOrderingPreservesSelectedAndRepeatedPositions() {
        byte[] package = Package(new[] { ("z", "z.xhtml", "application/xhtml+xml", ""),
            ("a", "a.xhtml", "application/xhtml+xml", "") }, "<itemref idref='a'/><itemref idref='z'/><itemref idref='a'/>",
            new[] { ("z.xhtml", Xhtml("<p>Z</p>")), ("a.xhtml", Xhtml("<p>A</p>")) });
        EpubDocument book = EpubDocument.Load(new MemoryStream(package), new EpubReadOptions {
            PreferSpineOrder = false, DeterministicOrder = false });
        Assert.Equal(new[] { "EPUB/z.xhtml", "EPUB/a.xhtml", "EPUB/a.xhtml" }, book.Chapters.Select(chapter => chapter.Path));
        Assert.Equal(new int?[] { 2, 1, 3 }, book.Chapters.Select(chapter => chapter.SpineIndex));
        Assert.True(book.ReadSummary.IsComplete);
    }

    [Fact]
    public void Load_RepeatedPositionsShareBoundedParsedPayloads() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg'><desc>" + new string('x', 4 * 1024 * 1024 - 100) +
            "</desc><text>Shared text</text></svg>";
        byte[] package = Package(new[] { ("s", "page.svg", "image/svg+xml", "") },
            string.Concat(Enumerable.Repeat("<itemref idref='s'/>", 500)), new[] { ("page.svg", svg) });
        EpubDocument book = EpubDocument.Load(new MemoryStream(package), new EpubReadOptions {
            MaxTotalUncompressedBytes = 5 * 1024 * 1024, IncludeRawHtml = true, MaxTotalRawHtmlBytes = 5 * 1024 * 1024 });
        Assert.Equal(500, book.Chapters.Count);
        Assert.All(book.Chapters, chapter => Assert.Same(book.Chapters[0].Text, chapter.Text));
        Assert.Equal(svg, book.Chapters[0].Html);
        Assert.All(book.Chapters.Skip(1), chapter => Assert.Null(chapter.Html));
        Assert.True(book.ReadSummary.IsComplete);
    }

    [Fact]
    public void Load_SvgForeignObjectBaseDoesNotChangeChapterBase() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg'><foreignObject><html xmlns='http://www.w3.org/1999/xhtml'>" +
            "<head><base href='https://example.org/elsewhere/'/></head><body/></html></foreignObject><image href='image.png'/></svg>";
        EpubDocument book = EpubDocument.Load(new MemoryStream(Package(new[] { ("s", "page.svg", "image/svg+xml", "") },
            "<itemref idref='s'/>", new[] { ("page.svg", svg) })));
        Assert.Null(Assert.Single(book.Chapters).BaseHref);
    }
}
