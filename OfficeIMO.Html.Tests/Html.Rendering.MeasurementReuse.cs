using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlMeasurementReuseTests {
    [Theory]
    [InlineData(PdfTextShapingMode.UnicodeScalar)]
    [InlineData(PdfTextShapingMode.LatinLigatures)]
    [InlineData(PdfTextShapingMode.OpenTypeLigatures)]
    public void RepeatedTextRetainsItsSizeAndFaceSpecificInlineAdvance(PdfTextShapingMode shaping) {
        byte[] regular = ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(
            new Dictionary<int, ushort> { ['A'] = 400, ['X'] = 400 });
        byte[] bold = ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(
            new Dictionary<int, ushort> { ['A'] = 900, ['X'] = 900 });
        const string html = "<style>body{margin:0;font-family:Measured}div{white-space:nowrap}"
            + "span:last-child{color:red}</style>"
            + "<div style='font-size:20px'><span>AA</span><span>X</span></div>"
            + "<div style='font-size:40px'><span>AA</span><span>X</span></div>"
            + "<div style='font-size:20px;font-weight:bold'><span>AA</span><span>X</span></div>"
            + "<div style='font-size:20px'><span>AA</span><span>X</span></div>";
        var options = new HtmlToPdfOptions {
            FontFamily = new PdfEmbeddedFontFamily("Measured", regular, bold),
            TextShapingMode = shaping,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };

        byte[] bytes = (HtmlConversionDocument.Parse(html).ToPdf(options)).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value is "A" or "X").ToArray();
        Assert.Equal("AAXAAXAAXAAX", string.Concat(letters.Select(letter => letter.Value)));
        double[] expected = { 12D, 24D, 27D, 12D };
        for (int row = 0; row < expected.Length; row++) {
            Assert.Equal(expected[row], letters[row * 3 + 2].StartBaseLine.X - letters[row * 3].StartBaseLine.X, 5);
        }
    }

    [Fact]
    public void ReusingCallerOptionsDoesNotCarryMeasurementsAcrossFontReplacement() {
        const string html = "<p style='font:20px Measured'><span>AA</span><span style='color:red'>X</span></p>";
        var document = HtmlConversionDocument.Parse(html);
        var options = new HtmlToPdfOptions {
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
        foreach (ushort advance in new ushort[] { 400, 800, 400 }) {
            options.FontFamily = new PdfEmbeddedFontFamily("Measured",
                ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(
                    new Dictionary<int, ushort> { ['A'] = advance, ['X'] = advance }));
            using var pdf = UglyToad.PdfPig.PdfDocument.Open((document.ToPdf(options)).ToBytes());
            var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value is "A" or "X").ToArray();
            Assert.Equal(3, letters.Length);
            Assert.Equal(advance * 20D * 0.75D * 2D / 1000D,
                letters[2].StartBaseLine.X - letters[0].StartBaseLine.X, 5);
        }
    }
}
