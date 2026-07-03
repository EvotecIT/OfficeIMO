using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {

    [Fact]
    public void SplitPages_ReadsExternalProducerPdfWithInheritedResourcesAndContentArrays() {
        byte[] pdf = BuildExternalTwoPagePdf();

        IReadOnlyList<byte[]> pages = PdfPageExtractor.SplitPages(pdf);

        Assert.Equal(2, pages.Count);
        Assert.Contains("External first page", Normalize(PdfTextExtractor.ExtractAllText(pages[0])), StringComparison.Ordinal);
        Assert.Contains("External second page", Normalize(PdfTextExtractor.ExtractAllText(pages[1])), StringComparison.Ordinal);
    }

    [Fact]
    public void Merge_ReordersExternalProducerPdfPagesAfterSplit() {
        byte[] pdf = BuildExternalTwoPagePdf();
        IReadOnlyList<byte[]> pages = PdfPageExtractor.SplitPages(pdf);

        byte[] merged = PdfMerger.Merge(pages[1], pages[0]);

        PdfDocumentInfo info = PdfInspector.Inspect(merged);
        string text = Normalize(PdfTextExtractor.ExtractAllText(merged));
        int secondPageIndex = text.IndexOf("External second page", StringComparison.Ordinal);
        int firstPageIndex = text.IndexOf("External first page", StringComparison.Ordinal);
        Assert.Equal(2, info.PageCount);
        Assert.NotEqual(-1, secondPageIndex);
        Assert.NotEqual(-1, firstPageIndex);
        Assert.True(secondPageIndex < firstPageIndex, text);
    }

    [Fact]
    public void SplitAndMerge_ReadExternalObjectStreamPageTree() {
        byte[] pdf = BuildExternalObjectStreamPdf(includeAcroForm: false);

        Assert.Contains("Object stream page", Normalize(PdfTextExtractor.ExtractAllText(pdf)), StringComparison.Ordinal);

        IReadOnlyList<byte[]> pages = PdfPageExtractor.SplitPages(pdf);
        Assert.Single(pages);
        Assert.Contains("Object stream page", Normalize(PdfTextExtractor.ExtractAllText(pages[0])), StringComparison.Ordinal);

        byte[] merged = PdfMerger.Merge(pdf, pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(merged);
        string mergedText = Normalize(PdfTextExtractor.ExtractAllText(merged));

        Assert.Equal(2, info.PageCount);
        Assert.Equal(2, CountOccurrences(mergedText, "Object stream page"));
    }

}
