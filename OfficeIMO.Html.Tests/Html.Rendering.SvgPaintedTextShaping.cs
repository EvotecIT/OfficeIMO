using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("A", false)]
    [InlineData("A\u0301", true)]
    public void HtmlSvgPaintedMarksRetainTheSharedImportLossInPng(string text, bool expectedLoss) {
        string html = "<style>@page{size:160px 70px;margin:0}html,body{margin:0}svg{display:block}</style>"
            + "<svg xmlns='http://www.w3.org/2000/svg' width='160' height='70'>"
            + "<defs><linearGradient id='ink'><stop stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient></defs>"
            + "<text x='8' y='40' font-family='Painted' font-size='22' fill='url(#ink)'>" + text + "</text></svg>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var options = new HtmlToPdfOptions {
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        };
        options.Fonts.Add("Painted", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphOutlines('A', 0x0301));

        OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(expectedLoss, png.Diagnostics.Any(diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.SvgContentUnsupported
            && diagnostic.LossKind != OfficeConversionLossKind.None));
        Assert.True(OfficeRasterImageDecoder.TryDecode(png.Bytes, out OfficeRasterImage? raster));
        Assert.Contains(Enumerable.Range(0, raster!.Height), y =>
            Enumerable.Range(0, raster.Width).Any(x => {
                OfficeColor pixel = raster.GetPixel(x, y);
                return pixel.A > 0 && (pixel.R != pixel.G || pixel.G != pixel.B);
            }));
        if (expectedLoss) {
            options.Policy.RequireNoLoss = true;
            OfficeImageExportPolicyException failure = Assert.Throws<OfficeImageExportPolicyException>(() =>
                document.ExportImage(OfficeImageExportFormat.Png, options));
            Assert.Contains(failure.Diagnostics, diagnostic =>
                diagnostic.Code == HtmlRenderDiagnosticCodes.SvgContentUnsupported
                && diagnostic.LossKind != OfficeConversionLossKind.None);
        }
    }
}
