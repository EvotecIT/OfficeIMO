using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderUserAgentStyleMode.Document, "", 200D)]
    [InlineData(HtmlRenderUserAgentStyleMode.Document, "content-box", 230D)]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser, "", 200D)]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser, "content-box", 230D)]
    public void HtmlTableGeometry_DeclaredWidthUsesTheDefaultOrAuthoredBox(HtmlRenderUserAgentStyleMode mode, string sizing, double width) {
        string html = "<body style='margin:0'><table id='allocated' style='width:200px;padding:10px;border:5px solid black;"
            + "margin:0;background:lime;" + (sizing.Length == 0 ? string.Empty : "box-sizing:" + sizing + ";")
            + "'><tr><td>Data</td></tr></table></body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            UserAgentStyles = mode,
            ViewportWidth = 300D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Equal(width, FindFlexShape(rendered, "table#allocated").Width, 3);
    }

    [Theory]
    [InlineData("display:flex;flex-direction:column", "", 180D)]
    [InlineData("display:flex;flex-direction:column", "content-box", 200D)]
    [InlineData("display:grid;grid-template-columns:200px", "", 180D)]
    [InlineData("display:grid;grid-template-columns:200px", "content-box", 200D)]
    public void HtmlTableGeometry_CrossAxisAllocationAppliesInsetsAndMaximumOnce(string containerStyle, string sizing, double width) {
        string html = "<body style='margin:0'><div style='width:200px;" + containerStyle + "'>"
            + "<table id='allocated' style='padding:0 10px;max-width:180px;min-width:0;margin:0;background:lime;"
            + (sizing.Length == 0 ? string.Empty : "box-sizing:" + sizing + ";")
            + "'><tr><td>Data</td></tr></table></div></body>";
        HtmlRenderDocument rendered = RenderFlex(html, 300D);

        Assert.Equal(width, FindFlexShape(rendered, "table#allocated").Width, 3);
    }
}
