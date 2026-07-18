using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageExtractorTests {
    [Fact]
    public void PdfPageRange_ParseReadsSingleAndInclusiveRanges() {
        Assert.Equal(PdfPageRange.From(3, 3), PdfPageRange.Parse("3"));
        Assert.Equal(PdfPageRange.From(1, 3), PdfPageRange.Parse(" 1-3 "));
        Assert.Equal(PdfPageRange.From(2, 4), PdfPageRange.Parse("2 .. 4"));
        Assert.Equal("2-4", PdfPageRange.Parse("2..4").ToString());
    }

    [Fact]
    public void PdfPageRange_ParseManyReadsWrapperFriendlyLists() {
        PdfPageRange[] ranges = PdfPageRange.ParseMany("1-2, 3; 2..3");

        Assert.Equal(new[] {
            PdfPageRange.From(1, 2),
            PdfPageRange.From(3, 3),
            PdfPageRange.From(2, 3)
        }, ranges);
    }

    [Fact]
    public void PdfPageRange_TryParseRejectsInvalidText() {
        Assert.True(PdfPageRange.TryParse("2-3", out PdfPageRange singleRange));
        Assert.Equal(PdfPageRange.From(2, 3), singleRange);

        Assert.True(PdfPageRange.TryParseMany("1,3-4", out PdfPageRange[] ranges));
        Assert.Equal(new[] { PdfPageRange.From(1, 1), PdfPageRange.From(3, 4) }, ranges);

        Assert.False(PdfPageRange.TryParse(null, out _));
        Assert.False(PdfPageRange.TryParse(" ", out _));
        Assert.False(PdfPageRange.TryParse("0", out _));
        Assert.False(PdfPageRange.TryParse("4-2", out _));
        Assert.False(PdfPageRange.TryParse("1-2-3", out _));
        Assert.False(PdfPageRange.TryParse("two", out _));
        Assert.False(PdfPageRange.TryParseMany(null, out _));
        Assert.False(PdfPageRange.TryParseMany("1,,2", out _));
    }

    [Fact]
    public void SplitPageRanges_AcceptsParsedRangeText() {
        byte[] source = BuildThreePagePdf();

        IReadOnlyList<byte[]> ranges = PdfPageExtractor.SplitPageRanges(source, PdfPageRange.ParseMany("1-2,3"));

        Assert.Equal(2, ranges.Count);
        string firstText = NormalizeExtractedText(PdfReadDocument.Open(ranges[0]).ExtractText());
        Assert.Contains("Firstpagemarker", firstText);
        Assert.Contains("Secondpagemarker", firstText);
        Assert.DoesNotContain("Thirdpagemarker", firstText);

        string secondText = NormalizeExtractedText(PdfReadDocument.Open(ranges[1]).ExtractText());
        Assert.DoesNotContain("Firstpagemarker", secondText);
        Assert.DoesNotContain("Secondpagemarker", secondText);
        Assert.Contains("Thirdpagemarker", secondText);
    }
}
