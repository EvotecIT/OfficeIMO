using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgReaderCrispEdgesTests {
    [Theory]
    [InlineData(1D, "1")]
    [InlineData(2D, "1")]
    [InlineData(1D, "0.5")]
    public void TransformedPixelStripsMatchSolidRectangleWithoutSeamsOrOverpainting(double scale, string opacity) {
        string svg = ReadStripFixture().Replace("fill-opacity=\"1\"", "fill-opacity=\"" + opacity + "\"");
        OfficeRasterImage actual = Render(svg, scale);
        OfficeRasterImage expected = Render(Wrap(
            "<rect x='1' y='1' width='2' height='2' fill='#1F2937' fill-opacity='" + opacity + "'/>"), scale);

        Assert.Equal(expected.GetPixels(), actual.GetPixels());
    }

    [Fact]
    public void SvgReexportPreservesCrispStripsAndNormalSiblingCoverage() {
        string svg = ReadStripFixture().Replace("</svg>",
            "<rect x='0.25' y='0.25' width='0.5' height='0.5' fill='black'/></svg>");
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing!, 1D, OfficeSvgSizeUnit.Pixel);

        OfficeRasterImage actual = Render(exported);
        OfficeRasterImage expected = Render(svg);

        Assert.Equal(OfficeColor.FromRgb(31, 41, 55), actual.GetPixel(1, 1));
        Assert.InRange(actual.GetPixel(0, 0).R, (byte)1, (byte)254);
        Assert.Equal(expected.GetPixels(), actual.GetPixels());
    }

    [Theory]
    [InlineData("style=\"shape-rendering:crispEdges\"")]
    [InlineData("class=\"strips\"")]
    [InlineData("shape-rendering=\"auto\" style=\"shape-rendering:crispEdges\"")]
    public void InlineAndStylesheetCrispEdgesInheritThroughNestedGroups(string declaration) {
        string svg = ReadStripFixture().Replace("shape-rendering=\"crispEdges\"", declaration)
            .Replace("<defs>", "<style>.strips { shape-rendering: crispEdges; }</style><defs>");
        OfficeRasterImage actual = Render(svg);

        Assert.Equal(OfficeColor.FromRgb(31, 41, 55), actual.GetPixel(1, 1));
        Assert.Equal(OfficeColor.FromRgb(31, 41, 55), actual.GetPixel(2, 2));
        Assert.Equal(OfficeColor.White, actual.GetPixel(0, 1));
    }

    [Theory]
    [InlineData("auto")]
    [InlineData("geometricPrecision")]
    [InlineData("initial")]
    public void NormalQualityOverridesCrispAncestorAndDoesNotLeakToSiblingShapes(string quality) {
        const string fractional = "<rect x='0.25' y='0.25' width='1.5' height='1.5' fill='black'/>";
        string svg = Wrap("<g shape-rendering='crispEdges'>"
            + "<rect x='3.25' y='0.25' width='0.5' height='1.5' fill='black'/>"
            + "<g style='shape-rendering:" + quality + "'>" + fractional + "</g></g>");
        OfficeRasterImage actual = Render(svg);
        OfficeRasterImage normal = Render(Wrap(fractional));

        Assert.Equal(OfficeColor.Black, actual.GetPixel(3, 0));
        for (int y = 0; y < 2; y++) for (int x = 0; x < 2; x++) {
            Assert.Equal(normal.GetPixel(x, y), actual.GetPixel(x, y));
            Assert.InRange(actual.GetPixel(x, y).R, (byte)1, (byte)254);
        }
    }

    [Fact]
    public void RectangularPathUsesTheSameCrispDeviceEdges() {
        OfficeRasterImage actual = Render(Wrap(
            "<path shape-rendering='crispEdges' transform='translate(0.25 0.25)' d='M0 0H1.5V1.5H0Z' fill='black'/>"));
        OfficeRasterImage expected = Render(Wrap("<rect width='2' height='2' fill='black'/>"));

        Assert.Equal(expected.GetPixels(), actual.GetPixels());
    }

    private static string ReadStripFixture() => File.ReadAllText(Path.Combine(RepositoryTestPaths.Find(),
        "OfficeIMO.Drawing.Tests", "TestAssets", "Svg", "crisp-edges-pixel-strips.svg"));

    private static string Wrap(string content) =>
        "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'>"
        + "<rect width='4' height='4' fill='white'/>" + content + "</svg>";

    private static OfficeRasterImage Render(string svg, double scale = 1D) {
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        return OfficeDrawingRasterRenderer.Render(drawing!, scale);
    }
}
