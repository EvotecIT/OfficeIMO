using System;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlFontTrackingTests {
    [Theory]
    [InlineData(10, 100)]
    [InlineData(15, 50)]
    [InlineData(20, 0)]
    public void EmbeddedPdfTrackingUsesCssSizeBeforeConversionToPoints(int size, int adjustment) {
        byte[] data = ManagedTextShapingTestAssets.AddTrackingTable(
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'B'),
            ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { 0D },
                new[] { new short[] { -100, 0 } }));
        HtmlConversionDocument source = HtmlConversionDocument.Parse("<style>@font-face{font-family:Tracked;src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(data) + "')}p{font-family:Tracked;font-size:" + size + "px}</style><p>AB</p>");
        var options = new HtmlToPdfOptions();
        options.PdfOptions.CompressContentStreams = false;
        byte[] pdf = source.ToPdfBytes(options);
        string content = Encoding.ASCII.GetString(pdf);
        Assert.Contains("<0002>" + (adjustment == 0 ? string.Empty : " " + adjustment) + "] TJ", content, StringComparison.Ordinal);
        Assert.Contains("AB", OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }
}
