using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlSvgWithExternalDtdUsesIntrinsicRatioForHeightOnlyImage() {
        const string svg = "<?xml version='1.0'?><!DOCTYPE svg PUBLIC '-//W3C//DTD SVG 1.1//EN' 'https://invalid.example.test/svg11.dtd'>"
            + "<svg xmlns='http://www.w3.org/2000/svg' width='120' height='100' viewBox='0 0 120 100'><rect width='120' height='100' fill='red'/></svg>";
        string html = "<style>body{margin:0}img{display:block;height:40px}</style>"
            + "<img src='data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)) + "'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 200D, ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile() });
        HtmlRenderDrawing image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        Assert.Equal(48D, image.Width, 3);
        Assert.Equal(40D, image.Height, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlSvgOpenPath_PreservesFillAndOpenStrokeAcrossExports(bool imageElement) {
        const string source = "<svg xmlns='http://www.w3.org/2000/svg' width='96' height='96'>"
            + "<path d='M8 8 H88 V88' fill='red' stroke='blue' stroke-width='4'/>"
            + "</svg>";
        string content = imageElement
            ? "<img src='data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(source)) + "'>"
            : source;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>html,body{margin:0}svg,img{display:block;width:96px;height:96px}</style>" + content);
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(1D), global::OfficeIMO.OfficeLength.Inches(1D)),
            ViewportWidth = 96D,
            ViewportHeight = 96D,
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.White,
            ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile()
        };

        OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, options);
        Assert.True(OfficeRasterImageDecoder.TryDecode(png.Bytes, out OfficeRasterImage? raster));
        OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, options);
        Assert.True(OfficeSvgDrawingReader.TryRead(svg.Bytes, out OfficeDrawing? svgDrawing, out int unsupported));
        Assert.Equal(0, unsupported);
        byte[] pdf = (document.ToPdf(new HtmlToPdfOptions(options))).ToBytes();
        Assert.Empty(OfficeIMO.Pdf.PdfImageExtractor.ExtractImages(pdf));
        OfficeDrawing pdfDrawing = OfficeIMO.Pdf.PdfReadDocument.Open(pdf).Pages[0].ToDrawing();

        Assert.All(new[] { raster!, OfficeDrawingRasterRenderer.Render(svgDrawing!, 1D, OfficeColor.White),
            OfficeDrawingRasterRenderer.Render(pdfDrawing, 4D / 3D, OfficeColor.White) }, image => {
            Assert.Equal(96, image.Width);
            Assert.Equal(96, image.Height);
            Assert.Equal(OfficeColor.Red, image.GetPixel(64, 32));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(87, 48));
            Assert.Equal(OfficeColor.Red, image.GetPixel(48, 47));
            Assert.Equal(OfficeColor.White, image.GetPixel(32, 64));
        });
    }

    [Fact]
    public void HtmlSvgNestedSymbolRetainsTheVisibleHeightThroughPdfEffectForms() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'>" +
            "<defs><symbol id='paint' viewBox='0 0 100 100'><rect width='100' height='100' fill='red'/></symbol>" +
            "<clipPath id='c'><rect width='10' height='20'/></clipPath></defs>" +
            "<use href='#paint' x='20' width='20' height='20' clip-path='url(#c)'/></svg>";
        string html = "<style>@page{size:40px 20px;margin:0}body{margin:0}img{display:block;width:40px;height:20px}</style>" +
            "<img src='data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)) + "'>";
        var options = new HtmlToPdfOptions { ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile() };
        byte[] pdf = (HtmlConversionDocument.Parse(html).ToPdf(options)).ToBytes();
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
        OfficeImageExportDiagnostic approximation = Assert.Single(png.Diagnostics);
        Assert.Equal("HtmlPdfDrawingEffectRasterized", approximation.Code);
        Assert.Equal(OfficeConversionLossKind.Approximation, approximation.LossKind);
        Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, approximation.Severity);
        HtmlRenderOptions strict = options.Clone();
        strict.Policy.RequireNoLoss = true;
        OfficeImageExportPolicyException strictFailure = Assert.Throws<OfficeImageExportPolicyException>(() =>
            document.ExportImage(OfficeImageExportFormat.Png, strict));
        Assert.Contains(strictFailure.Diagnostics, diagnostic =>
            diagnostic.Code == "HtmlPdfDrawingEffectRasterized"
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Equal(OfficeColor.FromRgb(246, 161, 153), reference.GetPixel(165, 50));
        Assert.Equal(OfficeColor.White, reference.GetPixel(225, 25));
        Assert.Equal(OfficeColor.FromRgb(25, 135, 84), reference.GetPixel(280, 180));
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_SVG_COMPOSITION_ARTIFACTS");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output);
            File.WriteAllText(Path.Combine(output, "source.svg"), source);
            File.WriteAllText(Path.Combine(output, "source.html"), html);
            File.WriteAllText(Path.Combine(output, "drawing.svg"), OfficeDrawingSvgExporter.ToSvg(drawing!));
            File.WriteAllBytes(Path.Combine(output, "composition.pdf"), (document.ToPdf(options)).ToBytes());
            File.WriteAllBytes(Path.Combine(output, "composition.png"), png.Bytes);
            File.WriteAllBytes(Path.Combine(output, "drawing.png"), OfficeDrawingRasterRenderer.ToPng(drawing!, 1D, OfficeColor.White));
            File.WriteAllText(Path.Combine(output, "diagnostics.json"), System.Text.Json.JsonSerializer.Serialize(png.Diagnostics));
        }
        foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Tiff, OfficeImageExportFormat.Webp }) {
            OfficeImageExportResult encoded = document.ExportImage(format, options);
            Assert.Equal(png.Diagnostics.Select(diagnostic => (diagnostic.Code, diagnostic.LossKind)),
                encoded.Diagnostics.Select(diagnostic => (diagnostic.Code, diagnostic.LossKind)));
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
