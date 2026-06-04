using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfTextExtractorPageTests {
    [Fact]
    public void ExtractTextByPage_ReadsFromCurrentStreamPosition() {
        byte[] pdf = BuildThreePagePdf();
        using var stream = BuildPrefixedStream(pdf);
        stream.Position = 5;

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPage(stream);

        Assert.Equal(3, pages.Count);
        Assert.Contains("Secondpagemarker", Normalize(pages[1]), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPageRanges_ReadsFromCurrentStreamPosition() {
        byte[] pdf = BuildThreePagePdf();
        using var stream = BuildPrefixedStream(pdf);
        stream.Position = 5;

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPageRanges(stream, PdfPageRange.ParseMany("2-3"));

        Assert.Equal(2, pages.Count);
        Assert.Contains("Secondpagemarker", Normalize(pages[0]), StringComparison.Ordinal);
        Assert.Contains("Thirdpagemarker", Normalize(pages[1]), StringComparison.Ordinal);
    }
}
