using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlSvgNestedSymbolRetainsTheVisibleHeightThroughPdfEffectForms() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'>" +
            "<defs><symbol id='paint' viewBox='0 0 100 100'><rect width='100' height='100' fill='red'/></symbol>" +
            "<clipPath id='c'><rect width='10' height='20'/></clipPath></defs>" +
            "<use href='#paint' x='20' width='20' height='20' clip-path='url(#c)'/></svg>";
        string html = "<style>@page{size:40px 20px;margin:0}body{margin:0}img{display:block;width:40px;height:20px}</style>" +
            "<img src='data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)) + "'>";
        var options = new HtmlToPdfOptions { ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile() };
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(OfficeIMO.Pdf.PdfPageImageRenderer.RenderPage(pdf));
        Assert.Equal(OfficeColor.Red, image.GetPixel(18, 8));
        Assert.Equal(OfficeColor.White, image.GetPixel(8, 8));
    }

    [Fact]
    public void HtmlSvgComposition_RetainsViewportClipsMasksAndAlphaAcrossLosslessFormats() {
        const string source = """
            <svg xmlns="http://www.w3.org/2000/svg" width="400" height="240" viewBox="0 0 400 240">
              <defs>
                <linearGradient id="paint" x1="0" y1="0" x2="1" y2="1">
                  <stop offset="0" stop-color="#e43"/><stop offset="1" stop-color="#26b"/>
                </linearGradient>
                <clipPath id="round"><circle cx="70" cy="50" r="42"/></clipPath>
                <mask id="fade" maskUnits="userSpaceOnUse" x="0" y="0" width="160" height="60">
                  <rect width="80" height="60" fill="white"/><rect x="80" width="80" height="60" fill="#808080"/>
                </mask>
                <symbol id="badge" viewBox="0 0 20 10"><rect width="20" height="10" fill="#198754"/></symbol>
              </defs>
              <rect width="400" height="240" fill="white"/>
              <svg x="20" y="20" width="160" height="100" viewBox="0 0 80 80" preserveAspectRatio="xMidYMid meet">
                <rect width="80" height="80" fill="#26b"/><rect x="40" width="80" height="80" fill="#e43" fill-opacity=".5"/>
              </svg>
              <g transform="translate(220 20)" clip-path="url(#round)"><rect width="140" height="100" fill="url(#paint)"/></g>
              <g transform="translate(20 150)" mask="url(#fade)"><rect width="160" height="60" fill="#e43"/></g>
              <use href="#badge" x="220" y="150" width="140" height="60"/>
            </svg>
            """;
        byte[] sourceBytes = Encoding.UTF8.GetBytes(source);
        Assert.True(OfficeSvgDrawingReader.TryRead(sourceBytes, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        string html = "<style>@page{size:400px 240px;margin:0}body{margin:0}img{display:block}</style>" +
            "<img width='400' height='240' src='data:image/svg+xml;base64," + Convert.ToBase64String(sourceBytes) + "'>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var options = new HtmlToPdfOptions { ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile() };
        OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, options);
        Assert.True(OfficeRasterImageDecoder.TryDecode(png.Bytes, out OfficeRasterImage? reference));
        Assert.Equal(400, reference!.Width);
        Assert.Equal(240, reference.Height);
        Assert.Empty(png.Diagnostics);
        Assert.Equal(OfficeColor.FromRgb(246, 161, 153), reference.GetPixel(165, 50));
        Assert.Equal(OfficeColor.White, reference.GetPixel(225, 25));
        Assert.Equal(OfficeColor.FromRgb(25, 135, 84), reference.GetPixel(280, 180));
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_SVG_COMPOSITION_ARTIFACTS");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output);
            File.WriteAllText(Path.Combine(output, "source.svg"), source);
            File.WriteAllText(Path.Combine(output, "source.html"), html);
            File.WriteAllText(Path.Combine(output, "drawing.svg"), OfficeDrawingSvgExporter.ToSvg(drawing!));
            File.WriteAllBytes(Path.Combine(output, "composition.pdf"), document.ToPdfBytes(options));
            File.WriteAllBytes(Path.Combine(output, "composition.png"), png.Bytes);
            File.WriteAllBytes(Path.Combine(output, "drawing.png"), OfficeDrawingRasterRenderer.ToPng(drawing!, 1D, OfficeColor.White));
            File.WriteAllText(Path.Combine(output, "diagnostics.json"), System.Text.Json.JsonSerializer.Serialize(png.Diagnostics));
        }
        foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Tiff, OfficeImageExportFormat.Webp }) {
            OfficeImageExportResult encoded = document.ExportImage(format, options);
            Assert.True(OfficeRasterImageDecoder.TryDecode(encoded.Bytes, out OfficeRasterImage? decoded));
            Assert.Equal(reference.Width, decoded!.Width);
            Assert.Equal(reference.Height, decoded.Height);
            for (int y = 0; y < reference.Height; y++) {
                for (int x = 0; x < reference.Width; x++) Assert.Equal(reference.GetPixel(x, y), decoded.GetPixel(x, y));
            }
            if (!string.IsNullOrWhiteSpace(output)) File.WriteAllBytes(Path.Combine(output, "composition" + encoded.FileExtension), encoded.Bytes);
        }
    }
}
