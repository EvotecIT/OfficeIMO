using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SvgFilterGraphsCarryNativePaintThroughInlineAndImageHtmlOutputs(bool imageResource) {
        const string svg = """
            <svg xmlns="http://www.w3.org/2000/svg" width="60" height="30" viewBox="0 0 60 30" style="display:block">
              <defs>
                <filter id="shadow" x="-30%" y="-30%" width="180%" height="180%">
                  <feOffset in="SourceAlpha" dx="3" dy="2" result="off"/>
                  <feComposite in="SourceGraphic" in2="off"/>
                </filter>
                <filter id="color">
                  <feColorMatrix values="0 0 1 0 0 0 1 0 0 0 1 0 0 0 0 0 0 0 1 0" result="matrix"/>
                  <feBlend in="SourceGraphic" in2="matrix" mode="multiply"/>
                </filter>
              </defs>
              <rect x="5" y="5" width="15" height="10" fill="red" filter="url(#shadow)"/>
              <rect x="35" y="5" width="15" height="10" fill="#c040a0" filter="url(#color)"/>
            </svg>
            """;
        string source = imageResource
            ? "<img src='data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg))
                + "' style='display:block;width:60px;height:30px'>"
            : svg;
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<body style='margin:0'>" + source + "</body>");
        var options = new HtmlRenderOptions { ViewportWidth = 60D, ViewportHeight = 30D, Margins = HtmlRenderMargins.All(0D) };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        Assert.DoesNotContain(rendered.Diagnostics, d => d.Code is HtmlRenderDiagnosticCodes.SvgContentUnsupported or HtmlRenderDiagnosticCodes.SvgRasterFallback);
        HtmlRenderDrawing visual = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        OfficeRasterImage scene = OfficeDrawingRasterRenderer.Render(visual.Drawing);
        Assert.Equal(OfficeColor.FromRgb(0, 0, 0), scene.GetPixel(21, 10));
        AssertSvgFilterPurple(scene.GetPixel(40, 10));
        Assert.True(OfficeRasterImageDecoder.TryDecode(document.ToPng(options), out OfficeRasterImage? png));
        AssertSvgFilterPurple(png!.GetPixel(40, 10));
        byte[] pdf = document.ToPdfDocumentResult(new HtmlToPdfOptions(options)).ToBytes();
        IReadOnlyList<PdfCore.PdfExtractedImage> images = PdfCore.PdfImageExtractor.ExtractImages(pdf);
        Assert.Contains(images, image => image.IsImageFile && image.MimeType == "image/png" &&
            OfficeRasterImageDecoder.TryDecode(image.Bytes, out OfficeRasterImage? decoded) && HasSvgFilterPurple(decoded!));
        Assert.Contains("data:image/png;base64,", document.ToSvg(options), StringComparison.Ordinal);
    }

    private static void AssertSvgFilterPurple(OfficeColor pixel) {
        Assert.InRange(pixel.R, (byte)118, (byte)120);
        Assert.InRange(pixel.G, (byte)8, (byte)10);
        Assert.InRange(pixel.B, (byte)118, (byte)120);
        Assert.Equal(255, pixel.A);
    }

    private static bool HasSvgFilterPurple(OfficeRasterImage image) {
        byte[] pixels = image.GetPixels();
        for (int i = 0; i < pixels.Length; i += 4)
            if (pixels[i] is >= 118 and <= 120 && pixels[i + 1] is >= 8 and <= 10 &&
                pixels[i + 2] is >= 118 and <= 120 && pixels[i + 3] == 255) return true;
        return false;
    }
}
