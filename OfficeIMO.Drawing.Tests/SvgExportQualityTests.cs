using System.Globalization;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class SvgExportQualityTests {
    [Fact]
    public void NearestNeighborKeepsCropFrameTransformAndOpacity() {
        var raster = new OfficeRasterImage(4, 2, OfficeColor.Red);
        var drawing = new OfficeDrawing(60, 50);
        drawing.AddImageWithInterpolation(OfficePngWriter.Encode(raster), "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(20, 18, 20, 12),
                new OfficeImageSourceCrop(.25, 0, .25, 0), rotationDegrees: 30,
                flipHorizontal: true), interpolate: false, opacity: .5);

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));
        XElement image = Assert.Single(svg.Descendants().Where(e => e.Name.LocalName == "image"));
        Assert.Equal("40", image.Attribute("width")?.Value);
        Assert.Equal("none", image.Attribute("preserveAspectRatio")?.Value);
        Assert.Contains(image.Ancestors(), e => e.Attribute("clip-path") != null && e.Attribute("transform") != null);
        Assert.Contains(image.Ancestors(), e => e.Attribute("opacity")?.Value == "0.5");
        Assert.Single(svg.Descendants().Where(e => e.Name.LocalName == "clipPath"));
    }

    [Theory]
    [InlineData(612, 792, 150D / 72D, "1275px", "1650px")]
    [InlineData(10.1, 20.1, 1, "11px", "21px")]
    [InlineData(.0000000001, .0000000001, 1, "1px", "1px")]
    public void PixelSurfaceUsesRasterRounding(double width, double height, double scale, string expectedWidth, string expectedHeight) {
        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(new OfficeDrawing(width, height), scale, OfficeSvgSizeUnit.Pixel));
        Assert.Equal(expectedWidth, svg.Attribute("width")?.Value);
        Assert.Equal(expectedHeight, svg.Attribute("height")?.Value);
    }

    [Theory]
    [InlineData(OfficeSoftMaskMode.Alpha, OfficeSoftMaskLuminosityStandard.Srgb)]
    [InlineData(OfficeSoftMaskMode.Luminosity, OfficeSoftMaskLuminosityStandard.Srgb)]
    [InlineData(OfficeSoftMaskMode.Luminosity, OfficeSoftMaskLuminosityStandard.PdfDeviceRgb)]
    public void SoftMaskTransfersFinalCoverageUsingAllSamples(OfficeSoftMaskMode mode, OfficeSoftMaskLuminosityStandard standard) {
        var content = new OfficeDrawing(12, 8);
        var maskDrawing = new OfficeDrawing(12, 8);
        OfficeShape maskShape = OfficeShape.Rectangle(6, 8);
        maskShape.FillColor = OfficeColor.FromRgba(255, 0, 0, 128);
        maskDrawing.AddShape(maskShape, 0, 0);
        var table = Enumerable.Range(0, 256).Select(i => (byte)(255 - i)).ToArray();
        var mask = new OfficeDrawingSoftMask(maskDrawing, mode, OfficeTransform.Identity, OfficeColor.FromRgba(0, 255, 0, 64), standard)
            .WithTransferTable(table);
        var drawing = new OfficeDrawing(12, 8).AddEffectDrawing(content, OfficeTransform.Identity, OfficeBlendMode.Normal, mask);

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing, 1, OfficeSvgSizeUnit.Pixel, null, "scene-"));
        XElement function = Assert.Single(svg.Descendants().Where(e => e.Name.LocalName == "feFuncR"));
        Assert.Equal(3, svg.Descendants().Count(e => e.Name.LocalName is "feFuncR" or "feFuncG" or "feFuncB"));
        Assert.Equal("discrete", function.Attribute("type")?.Value);
        double[] samples = function.Attribute("tableValues")!.Value.Split(' ').Select(v => double.Parse(v, CultureInfo.InvariantCulture)).ToArray();
        Assert.Equal(256, samples.Length);
        // Independent byte-domain lookup: each exact source coverage sample must
        // select the same transfer value as the raster owner's rounded lookup.
        for (int input = 0; input <= 255; input++) {
            int bin = Math.Min(255, (int)Math.Floor(input / 255D * 256));
            Assert.InRange(Math.Abs(table[input] / 255D - samples[bin]), 0D, .5e-8D);
        }
        XElement filter = Assert.Single(svg.Descendants().Where(e => e.Name.LocalName == "filter"));
        Assert.Equal("userSpaceOnUse", filter.Attribute("filterUnits")?.Value);
        Assert.Equal("12", filter.Attribute("width")?.Value);
        Assert.Equal("8", filter.Attribute("height")?.Value);
        Assert.Equal("sRGB", filter.Attribute("color-interpolation-filters")?.Value);
        XElement maskElement = Assert.Single(svg.Descendants().Where(e => e.Name.LocalName == "mask"));
        Assert.Equal("mask-type:luminance", maskElement.Attribute("style")?.Value);
        Assert.Contains(maskElement.Descendants(), e => e.Name.LocalName == "rect" && e.Attribute("fill")?.Value == "transparent");
        Assert.Contains(maskElement.Descendants(), e => e.Name.LocalName == "rect" && e.Attribute("fill")?.Value == "#00FF00"
            && e.Attribute("fill-opacity") != null);
        if (mode == OfficeSoftMaskMode.Luminosity) {
            XElement composite = Assert.Single(filter.Elements().Where(e => e.Name.LocalName == "feComposite"));
            Assert.Equal("SourceAlpha", composite.Attribute("in2")?.Value);
            Assert.Equal("in", composite.Attribute("operator")?.Value);
            string coefficients = filter.Elements().First(e => e.Name.LocalName == "feColorMatrix").Attribute("values")!.Value;
            Assert.Contains(standard == OfficeSoftMaskLuminosityStandard.PdfDeviceRgb ? "0.3 0.59 0.11" : "0.2126 0.7152 0.0722", coefficients);
            Assert.True(filter.Elements().Last().Name.LocalName == "feComponentTransfer");
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PositionedTextRetainsStrokeWithTransparentOrSuppressedFill(bool suppressFill) {
        var drawing = new OfficeDrawing(100, 40);
        drawing.AddPositionedTextWithNaturalAdvance("Outline", 2, 2, 90, 30,
            new OfficeFontInfo("Host font", 14), OfficeColor.Transparent, lineHeight: null);
        drawing.SetLastTextOutlineStroke(OfficeColor.FromRgba(0, 0, 255, 128), 2, suppressFill);

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));
        XElement text = Assert.Single(svg.Descendants().Where(e => e.Name.LocalName == "text"));
        Assert.Equal("Outline", text.Value);
        Assert.Equal("#0000FF", text.Attribute("stroke")?.Value);
        Assert.Equal("2", text.Attribute("stroke-width")?.Value);
        Assert.Equal("10", text.Attribute("stroke-miterlimit")?.Value);
        Assert.Equal(suppressFill ? "none" : "#000000", text.Attribute("fill")?.Value);
        if (!suppressFill) Assert.Equal("0", text.Attribute("fill-opacity")?.Value);
    }

    [Fact]
    public void PaintedGlyphRunsUseBoundedCorePathsAndKeepLogicalText() {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("FixtureFont", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddPositionedTextWithResolvedDirection("A", 2, 2, 90, 30,
            new OfficeFontInfo("FixtureFont", 14), OfficeColor.Red, null, 20, OfficeTextDirection.LeftToRight);
        drawing.MarkLastTextAsPaintedGlyphs();
        drawing.SetLastTextOutlineStroke(OfficeColor.Blue, 1.5, fillSuppressed: false);

        string markup = OfficeDrawingSvgExporter.ToSvg(drawing);
        var svg = XElement.Parse(markup);
        XElement path = Assert.Single(svg.Descendants().Where(e => e.Name.LocalName == "path"));
        Assert.NotEmpty(path.Attribute("d")!.Value);
        Assert.Equal("#FF0000", path.Attribute("fill")?.Value);
        Assert.Equal("#0000FF", path.Attribute("stroke")?.Value);
        Assert.Equal("1.5", path.Attribute("stroke-width")?.Value);
        Assert.Equal("A", path.Parent!.Attribute("aria-label")?.Value);
        Assert.DoesNotContain("data:font", markup);
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(markup), null, out OfficeDrawing? reopened, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Contains(OfficeDrawingRasterRenderer.Render(reopened!).GetPixels(), sample => sample != 0);
    }

    [Fact]
    public void BrowserTextRetainsReferencedFontsWithoutEmbeddingUnusedFaces() {
        byte[] used = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] unused = ManagedTextShapingTestAssets.CreateFont('B');
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("UsedFont", used);
        drawing.Fonts.Add("UnusedFont", unused);
        drawing.AddText("A", 2, 2, 90, 30, new OfficeFontInfo("UsedFont", 14), wrapText: false);

        string markup = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("<text", markup);
        Assert.Contains("data:font/ttf;base64," + Convert.ToBase64String(drawing.Fonts.Faces[0].Data), markup);
        Assert.DoesNotContain("UnusedFont", markup);
        Assert.DoesNotContain("data:font/ttf;base64," + Convert.ToBase64String(drawing.Fonts.Faces[1].Data), markup);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CommonSvgExportOffersGlyphPathsForOrdinaryText(bool asPaths) {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("FixtureFont", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30, new OfficeFontInfo("FixtureFont", 14), wrapText: false);

        OfficeImageExportResult result = drawing.ExportImage(OfficeImageExportFormat.Svg,
            new OfficeImageExportOptions { SvgTextAsPaths = asPaths });
        string markup = Encoding.UTF8.GetString(result.Bytes);
        var svg = XElement.Parse(markup);
        Assert.Equal(asPaths, svg.Descendants().Any(element => element.Name.LocalName == "path"));
        Assert.Equal(!asPaths, svg.Descendants().Any(element => element.Name.LocalName == "text"));
        Assert.Equal(!asPaths, markup.Contains("data:font"));
        if (asPaths) {
            XElement path = Assert.Single(svg.Descendants().Where(element => element.Name.LocalName == "path"));
            Assert.Equal("A", path.Parent!.Attribute("aria-label")?.Value);
            Assert.NotEmpty(path.Attribute("d")!.Value);
        }
        // Export settings belong to the detached export, not the caller's drawing.
        Assert.Contains("<text", OfficeDrawingSvgExporter.ToSvg(drawing));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GlyphPathsRetainPerCharacterFontFallback(bool painted) {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("Primary", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.Fonts.Add("Fallback", ManagedTextShapingTestAssets.CreateFont('A', 'B'));
        drawing.AddText("AB", 2, 2, 90, 30, new OfficeFontInfo("Primary, Fallback", 14), wrapText: false);
        if (painted) drawing.MarkLastTextAsPaintedGlyphs();
        Assert.Equal(2, drawing.Fonts.PlanFallbackRuns("AB", "Primary, Fallback").Count);
        string svg = Encoding.UTF8.GetString(drawing.ExportImage(OfficeImageExportFormat.Svg,
            new OfficeImageExportOptions { SvgTextAsPaths = !painted }).Bytes);
        Assert.Contains("<text", svg);
        Assert.DoesNotContain("<path", svg);
        Assert.Contains("data:font", svg);
    }

    [Fact]
    public void GlyphPathSettingSurvivesDrawingCloneAndNestedGroups() {
        var child = new OfficeDrawing(80, 30);
        child.Fonts.Add("FixtureFont", ManagedTextShapingTestAssets.CreateFont('A'));
        child.AddText("A", 1, 1, 70, 25, new OfficeFontInfo("FixtureFont", 14), wrapText: false);
        var drawing = new OfficeDrawing(100, 40).AddClippedDrawing(child, 2, 2, OfficeClipPath.Rectangle(80, 30));
        drawing.ApplyImageExportOptions(new OfficeImageExportOptions { SvgTextAsPaths = true });

        string markup = OfficeDrawingSvgExporter.ToSvg(drawing.Clone());
        var svg = XElement.Parse(markup);
        XElement path = Assert.Single(svg.Descendants().Where(element => element.Name.LocalName == "path"));
        Assert.Contains(path.Ancestors(), element => element.Attribute("clip-path") != null);
        Assert.DoesNotContain("data:font", markup);
    }

    [Fact]
    public void GlyphPathOptionRetainsUnsupportedWrappedTextAndItsFont() {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("FixtureFont", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30, new OfficeFontInfo("FixtureFont", 14), wrapText: true);

        string markup = Encoding.UTF8.GetString(drawing.ExportImage(OfficeImageExportFormat.Svg,
            new OfficeImageExportOptions { SvgTextAsPaths = true }).Bytes);
        Assert.Contains("<text", markup);
        Assert.Contains("data:font", markup);
    }

    [Fact]
    public void CommonSvgResultDimensionsAgreeWithRoundedRootSurface() {
        OfficeImageExportResult result = new OfficeDrawing(612, 792).ExportImage(OfficeImageExportFormat.Svg,
            new OfficeImageExportOptions { Scale = 150D / 72D });
        var svg = XElement.Parse(Encoding.UTF8.GetString(result.Bytes));
        Assert.Equal(1275, result.Width);
        Assert.Equal(1650, result.Height);
        Assert.Equal("1275px", svg.Attribute("width")?.Value);
        Assert.Equal("1650px", svg.Attribute("height")?.Value);
    }

    [Fact]
    public void EmbeddedImageStillObservesEncodedByteLimit() {
        var drawing = new OfficeDrawing(10, 10);
        drawing.AddImageWithInterpolation(OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Red)), "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 10, 10)), interpolate: false);
        Assert.Throws<OfficeImageExportBatchLimitException>(() => OfficeDrawingSvgExporter.ToSvgBytes(drawing,
            1, OfficeSvgSizeUnit.Pixel, null, null, maximumUtf8Bytes: 100, CancellationToken.None));
    }
}
