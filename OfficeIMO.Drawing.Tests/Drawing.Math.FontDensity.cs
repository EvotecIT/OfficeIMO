using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingMathFontDensityTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OutputDensityScalesTrackingGeometryWithoutChangingAuthoredSize(bool mathMetrics) {
        byte[] table = ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { -1D, 1D },
            new[] { new short[] { -120, -40 }, new short[] { -80, 40 } });
        var options = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Tracked", 15D), Padding = 0D, UseFontMathMetrics = mathMetrics };
        options.Fonts.Add("Tracked", ManagedTextShapingTestAssets.CreateTrackingFont(table));
        OfficeMathExpression expression = OfficeMath.Text("AB");
        options.Dpi = 72D;
        double normal = OfficeMathRenderer.Measure(expression, options).Width;
        options.Dpi = 144D;
        Assert.Equal(normal * 2D, OfficeMathRenderer.Measure(expression, options).Width, 6);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OutputDensityRetainsAuthoredOpticalSize(bool mathMetrics) {
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf"));
        OfficeMathRenderOptions Options(bool explicitSize) {
            var result = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Probe", 24D), Dpi = 144D, Padding = 0D, UseFontMathMetrics = mathMetrics };
            result.Fonts.FontVariationResolver = _ => explicitSize ? new Dictionary<string, float> { ["wght"] = 700F, ["opsz"] = 24F } : new Dictionary<string, float> { ["wght"] = 700F };
            result.Fonts.Add("Probe", data);
            return result;
        }
        OfficeMathExpression expression = OfficeMath.Text("Variable OfficeIMO");
        Assert.Equal(OfficeMathRenderer.Measure(expression, Options(true)).Width,
            OfficeMathRenderer.Measure(expression, Options(false)).Width, 6);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RasterDensityRetainsAuthoredOpticalSizeInMixedCopiedDrawing(bool copy) {
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf"));
        byte[] Paint(bool explicitSize) {
            var options = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Probe", 24D), Dpi = 144D, Padding = 0D };
            options.Fonts.FontVariationResolver = _ => explicitSize ? new Dictionary<string, float> { ["wght"] = 700F, ["opsz"] = 24F } : new Dictionary<string, float> { ["wght"] = 700F };
            options.Fonts.Add("Probe", data);
            var drawing = new OfficeDrawing(1000D, 160D);
            OfficeMathRenderer.AddToDrawing(drawing, OfficeMath.Text("Variable OfficeIMO"), 0D, 0D, options);
            drawing.AddText("Ordinary OfficeIMO", 0D, 100D, 600D, 50D, new OfficeFontInfo("Probe", 24D));
            if (copy) {
                var outer = new OfficeDrawing(1000D, 160D);
                outer.AddDrawing(drawing.Clone(), 0D, 0D);
                drawing = outer;
            }
            return OfficeDrawingRasterRenderer.ToPng(drawing);
        }
        Assert.Equal(Paint(true), Paint(false));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RasterDensityScalesTrackingPaintWithoutChangingAuthoredSize(bool copy) {
        byte[] table = ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { -1D, 1D },
            new[] { new short[] { -120, -40 }, new short[] { -80, 40 } });
        byte[] Paint(double density, double rasterScale) {
            var options = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Tracked", 15D), Padding = 0D, Dpi = density };
            options.Fonts.Add("Tracked", ManagedTextShapingTestAssets.CreateTrackingFont(table));
            var drawing = OfficeMathRenderer.Render(OfficeMath.Text("AB"), options);
            if (copy) {
                var outer = new OfficeDrawing(drawing.Width, drawing.Height);
                outer.AddDrawing(drawing.Clone(), 0D, 0D, new OfficeImageFrameTransform(180D, drawing.Width / 2D, drawing.Height / 2D));
                outer.ApplyColorTint(OfficeColor.Red);
                drawing = outer;
            }
            return OfficeDrawingRasterRenderer.ToPng(drawing, rasterScale);
        }
        Assert.Equal(Paint(72D, 2D), Paint(144D, 1D));
    }
    [Fact]
    public void SvgDensityRetainsSelectedVariableFontCoordinates() {
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf"));
        var options = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Probe", 24D), Dpi = 144D, Padding = 0D };
        options.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 700F };
        options.Fonts.Add("Probe", data);
        string svg = OfficeDrawingSvgExporter.ToSvg(OfficeMathRenderer.Render(OfficeMath.Text("Variable OfficeIMO"), options));
        var xml = System.Xml.Linq.XElement.Parse(svg);
        Assert.Contains(xml.DescendantsAndSelf(), e => e.Attribute("style")?.Value.Contains("\"opsz\" 24") == true &&
            e.Attribute("style")?.Value.Contains("\"wght\" 700") == true);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SvgDensityRetainsCoordinatesPerFallbackFace(bool variableFallback) {
        byte[] variable = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf"));
        var options = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Probe, Fallback", 24D), Dpi = 144D, Padding = 0D };
        options.Fonts.FontVariationResolver = request => request.FamilyName == "Probe"
            ? new Dictionary<string, float> { ["wght"] = 700F }
            : variableFallback ? new Dictionary<string, float> { ["wght"] = 300F, ["opsz"] = 36F } : null;
        options.Fonts.Add("Probe", variable, OfficeFontStyle.Regular,
            new OfficeFontUnicodeRangeSet(new[] { new OfficeFontUnicodeRange(65, 65) }));
        options.Fonts.Add("Fallback", variableFallback ? variable
            : ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(800, -200, 0, 200, 65, 66),
            OfficeFontStyle.Regular, variableFallback
                ? new OfficeFontUnicodeRangeSet(new[] { new OfficeFontUnicodeRange(66, 66) })
                : OfficeFontUnicodeRangeSet.All);
        var runs = options.Fonts.PlanFallbackRuns("AB", "Probe, Fallback");
        Assert.Equal(2, runs.Count);
        var xml = System.Xml.Linq.XElement.Parse(OfficeDrawingSvgExporter.ToSvg(
            OfficeMathRenderer.Render(OfficeMath.Text("AB"), options)));
        var a = Assert.Single(xml.Descendants(), e => e.Name.LocalName == "tspan" && e.Value == "A");
        Assert.Contains("\"opsz\" 24", a.Attribute("style")?.Value);
        Assert.Contains("\"wght\" 700", a.Attribute("style")?.Value);
        var b = Assert.Single(xml.Descendants(), e => e.Name.LocalName == "tspan" && e.Value == "B");
        if (variableFallback) {
            Assert.Contains("\"opsz\" 36", b.Attribute("style")?.Value);
            Assert.Contains("\"wght\" 300", b.Attribute("style")?.Value);
        } else Assert.Null(b.Attribute("style"));
    }

}
