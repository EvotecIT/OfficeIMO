using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfPdfLineSpacingRegressionTests {
    [Theory]
    [InlineData("direct-exact", false, -360, false, 36D)]
    [InlineData("direct-minimum", false, 360, false, 51.6D)]
    [InlineData("direct-multiple", false, 360, true, 54D)]
    [InlineData("inherited-exact", true, -360, false, 36D)]
    [InlineData("inherited-minimum", true, 360, false, 51.6D)]
    [InlineData("inherited-multiple", true, 360, true, 54D)]
    public void SignedSpacingRetainsItsUnitsAcrossMixedFontLines(
        string caseName, bool inherited, int spacing, bool multiple, double expectedSpan) {
        RtfDocument document = CreateMixedFontDocument();
        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        if (inherited) {
            RtfStyle baseStyle = document.AddStyle(1, "Base spacing");
            baseStyle.LineSpacingTwips = spacing;
            baseStyle.LineSpacingMultiple = multiple;
            document.AddStyle(2, "Inherited spacing").BasedOnStyleId = baseStyle.Id;
            paragraph.SetStyle(2);
        } else {
            paragraph.SetLineSpacing(spacing, multiple);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open((document.ToPdf(PortableOptions())).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var middle = Assert.Single(letters, letter => letter.Value == "B");
        var last = Assert.Single(letters, letter => letter.Value == "C");
        Assert.Equal(12D, first.FontSize, 3);
        Assert.Equal(24D, middle.FontSize, 3);
        Assert.Equal(12D, last.FontSize, 3);
        double span = first.StartBaseLine.Y - last.StartBaseLine.Y;
        Assert.True(Math.Abs(span - expectedSpan) < 0.01D,
            caseName + ": expected baseline span " + expectedSpan + ", actual " + span + ".");
        if (spacing < 0 && !multiple) {
            Assert.Equal(18D, first.StartBaseLine.Y - middle.StartBaseLine.Y, 3);
            Assert.Equal(18D, middle.StartBaseLine.Y - last.StartBaseLine.Y, 3);
        }
    }

    [Fact]
    public void DirectSpacingOverridesTheInheritedRuleAndValue() {
        RtfDocument document = CreateMixedFontDocument();
        RtfStyle style = document.AddStyle(1, "Large multiple spacing");
        style.LineSpacingTwips = 1200;
        style.LineSpacingMultiple = true;
        Assert.Single(document.Paragraphs).SetStyle(1).SetLineSpacing(-360, multiple: false);

        double[] baselines = ReadBaselines(document);

        Assert.Equal(18D, baselines[0] - baselines[1], 3);
        Assert.Equal(18D, baselines[1] - baselines[2], 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ZeroSpacingRetainsTheUnspecifiedPdfDefault(bool multiple) {
        RtfDocument unspecified = CreateMixedFontDocument();
        RtfDocument zero = CreateMixedFontDocument();
        Assert.Single(zero.Paragraphs).SetLineSpacing(0, multiple);

        double[] expected = ReadBaselines(unspecified);
        double[] actual = ReadBaselines(zero);

        for (int index = 0; index < expected.Length; index++) {
            Assert.Equal(expected[index], actual[index], 3);
        }
    }

    private static RtfDocument CreateMixedFontDocument() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("A").SetFontSize(12);
        paragraph.AddLineBreak();
        paragraph.AddText("B").SetFontSize(24);
        paragraph.AddLineBreak();
        paragraph.AddText("C").SetFontSize(12);
        return document;
    }

    private static double[] ReadBaselines(RtfDocument document) {
        using PdfPigDocument pdf = PdfPigDocument.Open((document.ToPdf(PortableOptions())).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters;
        return new[] { "A", "B", "C" }
            .Select(marker => Assert.Single(letters, letter => letter.Value == marker).StartBaseLine.Y)
            .ToArray();
    }

    private static RtfToPdfOptions PortableOptions() => new RtfToPdfOptions {
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        PdfOptions = new PdfCore.PdfCreateOptions {
            DefaultFont = PdfCore.PdfStandardFont.Helvetica,
            DefaultFontSize = 12
        }
    };
}
