using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlTextRetainsNumericWeightStretchAndObliqueAngleInItsDrawingAndSvg() {
        const string html = "<p style='font-weight:900;font-stretch:75%;font-style:oblique 20deg'>A<span style='font-weight:lighter'>B</span></p>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous };
        var rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        var runs = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>().ToArray();
        Assert.Contains(runs, run => run.Text.Contains("A") && run.Font.Face == new OfficeFontFaceDescriptor(900, 75, OfficeFontSlant.Oblique, 20));
        Assert.Contains(runs, run => run.Text.Contains("B") && run.Font.Face.Weight == 700);
        OfficeDrawingText[] drawingRuns = DrawingTestTraversal.Elements(rendered.Pages[0].CreateDrawing())
            .OfType<OfficeDrawingText>().ToArray();
        Assert.Contains(drawingRuns, run => run.Text.Contains("A") && run.Font.Face == new OfficeFontFaceDescriptor(900, 75, OfficeFontSlant.Oblique, 20));
        Assert.Contains(drawingRuns, run => run.Text.Contains("B") && run.Font.Face.Weight == 700);
        OfficeRasterImage svgRaster = RasterizeHtmlImageSvg(HtmlConversionDocument.Parse(html).ToSvg(options));
        Assert.All(runs, run => Assert.True(FontFallbackRunHasInk(svgRaster, run), run.Text));
    }
}
