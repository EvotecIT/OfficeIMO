using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageExtractorTests {
    [Fact]
    public void ExtractPages_CopiesSelectedPagesInRequestedOrder() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 3, 1);

        using var pdf = PdfPigDocument.Open(new MemoryStream(extracted));
        Assert.Equal(2, pdf.NumberOfPages);

        var read = PdfReadDocument.Open(extracted);
        Assert.Equal(2, read.Pages.Count);

        string text = NormalizeExtractedText(read.ExtractText());
        Assert.Contains("Thirdpagemarker", text);
        Assert.Contains("Firstpagemarker", text);
        Assert.DoesNotContain("Secondpagemarker", text);
        Assert.True(text.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < text.IndexOf("Firstpagemarker", StringComparison.Ordinal));

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal("Extraction sample", info.Metadata.Title);
        Assert.Equal("OfficeIMO", info.Metadata.Author);
        Assert.Equal(2, info.PageCount);
        Assert.Equal(300, info.Pages[0].Width);
        Assert.Equal(500, info.Pages[0].Height);
        Assert.Equal(595, info.Pages[1].Width);
        Assert.Equal(842, info.Pages[1].Height);
    }

    [Fact]
    public void ExtractPages_ClonesDuplicateSelectionsInRequestedOrder() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 3, 3, 1);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal(3, info.PageCount);
        Assert.Equal(300, info.Pages[0].Width);
        Assert.Equal(500, info.Pages[0].Height);
        Assert.Equal(300, info.Pages[1].Width);
        Assert.Equal(500, info.Pages[1].Height);
        Assert.Equal(595, info.Pages[2].Width);
        Assert.Equal(842, info.Pages[2].Height);

        string text = NormalizeExtractedText(PdfReadDocument.Open(extracted).ExtractText());
        Assert.Equal(2, CountOccurrences(text, "Thirdpagemarker"));
        Assert.Equal(1, CountOccurrences(text, "Firstpagemarker"));
        Assert.DoesNotContain("Secondpagemarker", text);
        AssertContainsInOrder(text, "Thirdpagemarker", "Thirdpagemarker", "Firstpagemarker");
    }

    [Fact]
    public void ExtractPageRange_CopiesInclusiveRange() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPageRange(source, 2, 3);

        var read = PdfReadDocument.Open(extracted);
        string text = NormalizeExtractedText(read.ExtractText());

        Assert.Equal(2, read.Pages.Count);
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }

    [Fact]
    public void ExtractPageRange_AcceptsPdfPageRange() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPageRange(source, PdfPageRange.From(2, 3));

        var read = PdfReadDocument.Open(extracted);
        string text = NormalizeExtractedText(read.ExtractText());

        Assert.Equal(2, read.Pages.Count);
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }

    [Fact]
    public void ExtractPageRanges_CombinesParsedRangesInCallerOrder() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPageRanges(source, PdfPageRange.ParseMany("3,1-2"));

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal(3, info.PageCount);
        Assert.Equal(300, info.Pages[0].Width);
        Assert.Equal(500, info.Pages[0].Height);
        Assert.Equal(595, info.Pages[1].Width);
        Assert.Equal(842, info.Pages[1].Height);
        Assert.Equal(612, info.Pages[2].Width);
        Assert.Equal(792, info.Pages[2].Height);

        string text = NormalizeExtractedText(PdfReadDocument.Open(extracted).ExtractText());
        AssertContainsInOrder(text, "Thirdpagemarker", "Firstpagemarker", "Secondpagemarker");
    }

    [Fact]
    public void ExtractPageRanges_PreservesDuplicateAndOverlappingRanges() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPageRanges(source, PdfPageRange.ParseMany("2,2-3"));

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal(3, info.PageCount);

        string text = NormalizeExtractedText(PdfReadDocument.Open(extracted).ExtractText());
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.Equal(2, CountOccurrences(text, "Secondpagemarker"));
        Assert.Equal(1, CountOccurrences(text, "Thirdpagemarker"));
        AssertContainsInOrder(text, "Secondpagemarker", "Secondpagemarker", "Thirdpagemarker");
    }
}
