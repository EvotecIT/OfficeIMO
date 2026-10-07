using System;
using System.Linq;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlFontTrackingTests {
    [Theory]
    [InlineData(10, 100, PdfTextShapingMode.UnicodeScalar)]
    [InlineData(15, 50, PdfTextShapingMode.UnicodeScalar)]
    [InlineData(20, 0, PdfTextShapingMode.UnicodeScalar)]
    [InlineData(10, 100, PdfTextShapingMode.OpenTypeLigatures)]
    [InlineData(15, 50, PdfTextShapingMode.OpenTypeLigatures)]
    [InlineData(20, 0, PdfTextShapingMode.OpenTypeLigatures)]
    public void EmbeddedPdfTrackingUsesCssSizeBeforeConversionToPoints(int size, int adjustment, PdfTextShapingMode shapingMode) {
        byte[] data = ManagedTextShapingTestAssets.AddTrackingTable(
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'B'),
            ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { 0D },
                new[] { new short[] { -100, 0 } }));
        HtmlConversionDocument source = HtmlConversionDocument.Parse("<style>@font-face{font-family:Tracked;src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(data) + "')}p{font-family:Tracked;font-size:" + size + "px}</style><p>AB</p>");
        var options = new HtmlToPdfOptions();
        options.PdfOptions.CompressContentStreams = false;
        options.PdfOptions.TextShapingMode = shapingMode;
        byte[] pdf = source.ToPdfBytes(options);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(pdf);
        var letters = independent.GetPage(1).Letters.ToArray();
        Assert.Equal(new[] { "A", "B" }, letters.Select(letter => letter.Value));
        double expectedAdvance = (500D - adjustment) * size * 0.75D / 1000D;
        Assert.Equal(expectedAdvance, letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, precision: 5);
        Assert.Contains("AB", OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }
}
