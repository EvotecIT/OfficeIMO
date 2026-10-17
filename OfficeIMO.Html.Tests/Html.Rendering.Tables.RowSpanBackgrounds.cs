using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("collapse")]
    [InlineData("separate")]
    public void HtmlTables_LaterRowBackgroundStaysBelowOpaqueRowSpan(string borderModel) {
        HtmlRenderDocument rendered = RenderRowSpanBackgrounds(borderModel);
        Assert.Single(rendered.Pages);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.Red, image.GetPixel(30, 110));
        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), image.GetPixel(200, 110));
        Assert.Equal(2, RowSpanBackgroundCell(rendered, "td#span-1").RowSpan);
        AssertRowSpanBackgroundLabels(rendered, "SPAN-1", "FIRST-1", "SECOND-1");
    }

    [Theory]
    [InlineData("collapse")]
    [InlineData("separate")]
    public void HtmlTables_RowSpanLayersStayScopedAroundNestedTableAndOpacity(string borderModel) {
        HtmlRenderDocument reference = RenderRowSpanBackgrounds(borderModel, nested: true, laterColor: "transparent");
        HtmlRenderDocument colored = RenderRowSpanBackgrounds(borderModel, nested: true);
        Assert.Single(reference.Pages);
        Assert.Single(colored.Pages);
        HtmlRenderSemanticGroup span = RowSpanBackgroundCell(colored, "td#span-1");
        HtmlRenderSemanticGroup referenceSpan = RowSpanBackgroundCell(reference, "td#span-1");
        Assert.Equal(referenceSpan.X, span.X, 3);
        Assert.Equal(referenceSpan.Y, span.Y, 3);
        Assert.Equal(referenceSpan.Width, span.Width, 3);
        Assert.Equal(referenceSpan.Height, span.Height, 3);
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(reference.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(colored.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        int sideX = (int)(span.X + span.Width - 8D);
        int lowerY = (int)(span.Y + span.Height - 20D);
        int innerX = (int)(span.X + 20D);
        Assert.True(expected.GetPixel(sideX, lowerY).R > expected.GetPixel(sideX, lowerY).B);
        Assert.True(expected.GetPixel(innerX, lowerY).B > expected.GetPixel(innerX, lowerY).R);
        Assert.Equal(expected.GetPixel(sideX, lowerY), actual.GetPixel(sideX, lowerY));
        Assert.Equal(expected.GetPixel(innerX, lowerY), actual.GetPixel(innerX, lowerY));
        Assert.Equal(expected.GetPixel(innerX, (int)span.Y + 20), actual.GetPixel(innerX, (int)span.Y + 20));
        AssertRowSpanBackgroundLabels(colored, "INNER-A", "INNER-B", "FIRST-1", "SECOND-1");
    }

    [Theory]
    [InlineData("collapse")]
    [InlineData("separate")]
    public void HtmlTables_RowSpanBackgroundsSurviveGroupBreaksAndOwnRepeatedSections(string borderModel) {
        HtmlRenderDocument rendered = RenderRowSpanBackgrounds(borderModel, repeated: true);
        Assert.Equal(2, rendered.Pages.Count);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        Assert.Equal(2, text.Count(item => item.Text == "HEADER"));
        Assert.Equal(2, text.Count(item => item.Text == "FOOTER"));
        AssertRowSpanBackgroundLabels(rendered, "SPAN-1", "FIRST-1", "SECOND-1", "SPAN-2", "FIRST-2", "SECOND-2");
        for (int page = 0; page < rendered.Pages.Count; page++) {
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(rendered.Pages[page].CreateDrawing(), 1D, OfficeColor.White);
            Assert.Equal(OfficeColor.Red, image.GetPixel(30, 140));
            Assert.Equal(OfficeColor.FromRgb(0x23, 0x39, 0x5d), image.GetPixel(250, 30));
            HtmlRenderSemanticGroup footer = EnumerateRenderVisuals(rendered.Pages[page].Scene)
                .OfType<HtmlRenderSemanticGroup>().First(group => group.Role == HtmlRenderSemanticGroupRole.TableCell
                    && group.Source == "td#footer");
            Assert.Equal(OfficeColor.FromRgb(0xdd, 0xdd, 0xdd), image.GetPixel(250, (int)footer.Y + 5));
        }
        Assert.Equal(2, RowSpanBackgroundCell(rendered, "td#span-1").RowSpan);
        Assert.Equal(2, RowSpanBackgroundCell(rendered, "td#span-2").RowSpan);
    }

    private static HtmlRenderSemanticGroup RowSpanBackgroundCell(HtmlRenderDocument rendered, string source) =>
        rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderSemanticGroup>().First(group => group.Role == HtmlRenderSemanticGroupRole.TableCell
                && group.Source == source);

    private static void AssertRowSpanBackgroundLabels(HtmlRenderDocument rendered, params string[] labels) {
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        foreach (string label in labels) Assert.Single(text, item => item.Text == label);
    }

    private static HtmlRenderDocument RenderRowSpanBackgrounds(
        string borderModel, bool nested = false, bool repeated = false, string laterColor = "lime") {
        string inner = "<table class='inner'><tbody><tr><td>INNER-A</td></tr><tr><td>INNER-B</td></tr></tbody></table>";
        string Pair(int index) => "<tbody style='background:cyan'><tr><td id='span-" + index
            + "' class='span' rowspan='2'>" + (nested ? inner : "SPAN-" + index)
            + "</td><td>FIRST-" + index + "</td></tr><tr class='later'><td>SECOND-" + index + "</td></tr></tbody>";
        string html = "<style>@page{size:400px 320px;margin:20px}html,body{margin:0;background:white;font:14px/18px Arial}"
            + "table{width:300px;table-layout:fixed;border-collapse:" + borderModel + ";border-spacing:0;background:#6bb7e6}"
            + "tbody>tr{height:60px;background:#f5d47a}td,th{border:2px solid #172033;padding:4px}"
            + "tr.later{background:" + laterColor + "}td.span{background:red}.inner{width:100px}.inner td{background:blue}"
            + ".scope{opacity:.5}thead tr{height:28px}thead{background:#23395d;color:white}tfoot{background:#ddd}"
            + "</style>" + (nested ? "<div class='scope'>" : string.Empty) + "<table class='outer'>"
            + (repeated ? "<thead><tr><th colspan='2'>HEADER</th></tr></thead>" : string.Empty)
            + Pair(1) + (repeated ? Pair(2) + "<tfoot><tr><td id='footer' colspan='2'>FOOTER</td></tr></tfoot>" : string.Empty)
            + "</table>" + (nested ? "</div>" : string.Empty);
        return HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true,
            ViewportWidth = 400D, ViewportHeight = 320D, Margins = HtmlRenderMargins.All(0D)
        });
    }
}
