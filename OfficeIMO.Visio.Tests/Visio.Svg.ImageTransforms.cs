using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioSvgImageTransformTests {
    [Theory]
    [InlineData(0D, true, false, false)]
    [InlineData(0D, false, true, false)]
    [InlineData(0D, true, false, true)]
    [InlineData(0D, false, true, true)]
    [InlineData(90D, true, false, true)]
    [InlineData(90D, false, true, true)]
    [InlineData(90D, true, true, true)]
    [InlineData(0D, false, false, true)]
    public void SvgImagePreviewPreservesPlacementAndPixelsThroughCroppingAndReflection(
        double rotationDegrees, bool flipHorizontal, bool flipVertical, bool cropped) {
        OfficeRasterImage source = CreateQuadrantImage();
        var projection = new OfficeImageProjection(
            new OfficeImagePlacement(40, 60, 80, 80),
            cropped ? new OfficeImageSourceCrop(.25, .25, .25, .25) : default,
            rotationDegrees,
            flipHorizontal: flipHorizontal,
            flipVertical: flipVertical);
        var expected = new OfficeRasterImage(200, 200, OfficeColor.Transparent);
        new OfficeRasterCanvas(expected).DrawImage(source, projection);

        const string svgNamespace = "http://www.w3.org/2000/svg";
        var svg = new StringBuilder();
        using (XmlWriter writer = XmlWriter.Create(svg, new XmlWriterSettings { OmitXmlDeclaration = true })) {
            writer.WriteStartElement("svg", svgNamespace);
            writer.WriteAttributeString("width", "200");
            writer.WriteAttributeString("height", "200");
            OfficeSvgImageRenderer.WriteImage(writer, svgNamespace,
                OfficeSvgImageRenderer.CreateDataUri("image/png", OfficePngWriter.Encode(source)),
                projection, preserveAspectRatio: "none", clipPathId: "crop");
            writer.WriteEndElement();
        }

        var diagnostics = new List<OfficeImageExportDiagnostic>();
        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg.ToString()), null, null, null, null, null,
            diagnostics, "transformed-image.svg", default, out OfficeRasterImage? image));
        OfficeRasterImage actual = Assert.IsType<OfficeRasterImage>(image);
        foreach ((int x, int y) in new[] { (60, 80), (100, 80), (60, 120), (100, 120),
            (30, 100), (80, 40), (130, 100), (80, 160) }) {
            Assert.Equal(expected.GetPixel(x, y), actual.GetPixel(x, y));
        }
        Assert.Empty(diagnostics);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ShearedImagePreservesSourceQuadrantsAndParallelogramBounds(bool horizontal) {
        string svg = CreateImageSvg(horizontal ? "skewX(45)" : "skewY(45)");
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        OfficeRasterImage actual = Rasterize(svg, diagnostics);
        (int X, int Y)[] samples = horizontal
            ? new[] { (60, 30), (80, 30), (80, 50), (100, 50) }
            : new[] { (30, 60), (50, 80), (30, 80), (50, 100) };
        OfficeColor[] colors = { OfficeColor.Red, OfficeColor.Green, OfficeColor.Blue, OfficeColor.Yellow };
        for (int index = 0; index < samples.Length; index++) {
            Assert.Equal(colors[index], actual.GetPixel(samples[index].X, samples[index].Y));
        }
        (int X, int Y)[] outside = horizontal ? new[] { (45, 50), (110, 30) } : new[] { (50, 45), (30, 110) };
        foreach ((int x, int y) in outside) {
            Assert.Equal(OfficeColor.Transparent, actual.GetPixel(x, y));
        }
        Assert.Empty(diagnostics);
    }

    [Fact]
    public void DegenerateImageTransformReportsLossWithoutPaintingItsBoundingRectangle() {
        string svg = CreateImageSvg("matrix(1 1 1 1 0 0)", whiteBackground: true);
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        OfficeRasterImage actual = Rasterize(svg, diagnostics);

        Assert.Equal(OfficeColor.White, actual.GetPixel(60, 60));
        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    private static string CreateImageSvg(string transform, bool whiteBackground = false) {
        string href = OfficeSvgImageRenderer.CreateDataUri("image/png", OfficePngWriter.Encode(CreateQuadrantImage()));
        return "<svg xmlns='http://www.w3.org/2000/svg' width='140' height='140'>" +
               (whiteBackground ? "<rect width='140' height='140' fill='white'/>" : "") +
               "<image x='20' y='20' width='40' height='40' transform='" + transform +
               "' preserveAspectRatio='none' href='" + href + "'/></svg>";
    }

    private static OfficeRasterImage Rasterize(string svg, List<OfficeImageExportDiagnostic> diagnostics) {
        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "affine-image.svg", default, out OfficeRasterImage? image));
        return Assert.IsType<OfficeRasterImage>(image);
    }

    private static OfficeRasterImage CreateQuadrantImage() {
        var image = new OfficeRasterImage(16, 16, OfficeColor.Transparent);
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                image.SetPixel(x, y, y < 8
                    ? (x < 8 ? OfficeColor.Red : OfficeColor.Green)
                    : (x < 8 ? OfficeColor.Blue : OfficeColor.Yellow));
            }
        }
        return image;
    }
}
