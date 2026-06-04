using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfTextExtractorPageTests {
    [Fact]
    public void ExtractTextByPage_ReturnsOneTextEntryPerPage() {
        byte[] pdf = BuildThreePagePdf();

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPage(pdf);

        Assert.Equal(3, pages.Count);
        Assert.Contains("Firstpagemarker", Normalize(pages[0]), StringComparison.Ordinal);
        Assert.Contains("Secondpagemarker", Normalize(pages[1]), StringComparison.Ordinal);
        Assert.Contains("Thirdpagemarker", Normalize(pages[2]), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_NullLayoutOptionsUseNoOptionsPath() {
        byte[] pdf = BuildThreePagePdf();
        IReadOnlyList<string> expected = PdfTextExtractor.ExtractTextByPage(pdf);

        IReadOnlyList<string> bytePages = PdfTextExtractor.ExtractTextByPage(pdf, (PdfTextLayoutOptions?)null);
        using var stream = new MemoryStream(pdf);
        IReadOnlyList<string> streamPages = PdfTextExtractor.ExtractTextByPage(stream, (PdfTextLayoutOptions?)null);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");

        try {
            File.WriteAllBytes(path, pdf);
            IReadOnlyList<string> pathPages = PdfTextExtractor.ExtractTextByPage(path, (PdfTextLayoutOptions?)null);

            Assert.Equal(expected, bytePages);
            Assert.Equal(expected, streamPages);
            Assert.Equal(expected, pathPages);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ExtractTextByPageRanges_ReturnsParsedRangesInCallerOrderWithRepeatedPages() {
        byte[] pdf = BuildThreePagePdf();

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPageRanges(pdf, PdfPageRange.ParseMany("3,1-2,2"));

        Assert.Equal(4, pages.Count);
        Assert.Contains("Thirdpagemarker", Normalize(pages[0]), StringComparison.Ordinal);
        Assert.Contains("Firstpagemarker", Normalize(pages[1]), StringComparison.Ordinal);
        Assert.Contains("Secondpagemarker", Normalize(pages[2]), StringComparison.Ordinal);
        Assert.Contains("Secondpagemarker", Normalize(pages[3]), StringComparison.Ordinal);
    }
}
