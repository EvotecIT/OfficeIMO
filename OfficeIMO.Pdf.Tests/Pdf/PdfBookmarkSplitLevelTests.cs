using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfBookmarkSplitLevelTests {
    [Fact]
    public void LevelSectionsRetainNestedPagesAndStopAtOuterSections() {
        byte[] source = PdfDocument.Create(new PdfCreateOptions { CreateOutlineFromHeadings = true })
            .H1("Chapter One").Paragraph(p => p.Text("Opening body"))
            .PageBreak().H2("Nested section").Paragraph(p => p.Text("Nested body"))
            .PageBreak().H1("Chapter Two").Paragraph(p => p.Text("Next chapter body"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(new System.IO.MemoryStream(source));
        Assert.Equal(3, document.Pages.BookmarkPageRanges().Count);
        var chapters = document.Pages.BookmarkPageRanges(1);
        Assert.Equal(2, chapters.Count);
        Assert.Equal(PdfPageRange.From(1, 2), chapters[0].PageRange);
        Assert.Equal(PdfPageRange.From(3, 3), chapters[1].PageRange);
        var nested = Assert.Single(document.Pages.BookmarkPageRanges(2, CancellationToken.None, "Nested section"));
        Assert.Equal(PdfPageRange.From(2, 2), nested.PageRange);
        using PdfDocument extracted = document.Pages.Extract(chapters[0].PageRange);
        Assert.Contains("Nested body", extracted.Read().Text, System.StringComparison.Ordinal);
        Assert.DoesNotContain("Next chapter body", extracted.Read().Text, System.StringComparison.Ordinal);
        Assert.Throws<System.ArgumentException>(() => document.Pages.BookmarkPageRanges(1, "Nested section"));
        Assert.Throws<System.OperationCanceledException>(() => document.Pages.BookmarkPageRanges(1, new CancellationToken(true)));
    }
}
