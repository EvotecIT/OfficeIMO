using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlTextRetainsNumericWeightStretchAndObliqueAngleInItsDrawingAndSvg() {
        const string html = "<p style='font-weight:900;font-stretch:75%;font-style:oblique 20deg'>A<span style='font-weight:lighter'>B</span></p>";
        var rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        var runs = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>().ToArray();
        Assert.Contains(runs, run => run.Text.Contains("A") && run.Font.Face == new OfficeFontFaceDescriptor(900, 75, OfficeFontSlant.Oblique, 20));
        Assert.Contains(runs, run => run.Text.Contains("B") && run.Font.Face.Weight == 700);
        string svg = HtmlConversionDocument.Parse(html).ToSvg();
        Assert.Contains("font-weight=\"900\"", svg);
        Assert.Contains("font-stretch=\"75%\"", svg);
        Assert.Contains("font-style=\"oblique 20deg\"", svg);
    }
}
