using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("auto", "", 4, 80D)]
    [InlineData("auto", "max-height:60px", 3, 60D)]
    [InlineData("auto", "max-height:100px", 4, 80D)]
    [InlineData("auto", "min-height:100px", 4, 100D)]
    [InlineData("balance", "min-height:60px", 2, 60D)]
    public void HtmlColumns_AutoHeightSeparatesSequentialFillFromBalancedContent(
        string fill, string bounds, int blocksPerColumn, double expectedHeight) {
        string html = ColumnAutoSizingSource("column-fill:" + fill + ";" + bounds);

        HtmlRenderDocument rendered = RenderColumns(html, 340D);

        for (int index = 0; index < 4; index++) {
            HtmlRenderShape shape = FindColumnShape(rendered, "div#sizing-" + index);
            Assert.Equal(160D * (index / blocksPerColumn), shape.X, 3);
            Assert.Equal(20D * (index % blocksPerColumn), shape.Y, 3);
            Assert.Equal(20D, shape.Height, 3);
        }
        Assert.Equal(expectedHeight, FindColumnShape(rendered, "section#sizing-columns").Height, 3);
        Assert.Equal(expectedHeight, FindColumnShape(rendered, "div#sizing-after").Y, 3);
    }

    [Theory]
    [InlineData("break-before:column")]
    [InlineData("break-after:column")]
    [InlineData("break-before:avoid-column")]
    [InlineData("break-after:avoid-column")]
    [InlineData("break-inside:avoid-column")]
    public void HtmlColumns_UnsupportedColumnBreaksPreserveSourceAndReportTypedLoss(string declaration) {
        string html = ColumnAutoSizingSource("column-fill:auto", thirdBlockStyle: declaration);
        foreach (HtmlRenderMode mode in new[] { HtmlRenderMode.Continuous, HtmlRenderMode.Paged }) {
            HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html),
                new HtmlRenderOptions { Mode = mode, ViewportWidth = 340D, Margins = HtmlRenderMargins.All(0D) });

            for (int index = 0; index < 4; index++) {
                Assert.Equal(20D, FindColumnShape(rendered, "div#sizing-" + index).Height, 3);
            }
            HtmlDiagnostic[] diagnostics = rendered.Diagnostics.Where(
                item => item.Code == HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported).ToArray();
            Assert.NotEmpty(diagnostics);
            Assert.All(diagnostics, diagnostic => {
                Assert.Equal("div#sizing-2", diagnostic.Source);
                Assert.Contains(declaration.Replace(':', '='), diagnostic.Detail);
            });
            Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
        }
    }

    [Theory]
    [InlineData("content-box", 3, 160D, 324D, 84D)]
    [InlineData("border-box", 1, 148D, 300D, 60D)]
    public void HtmlColumns_AutoMaximumHeightUsesTheDeclaredBoxSizingForColumnCapacity(
        string boxSizing, int blocksPerColumn, double columnAdvance, double expectedWidth, double expectedHeight) {
        string html = ColumnAutoSizingSource("column-fill:auto;max-height:60px;box-sizing:" + boxSizing
            + ";padding:10px;border:2px solid black");

        HtmlRenderDocument rendered = RenderColumns(html, 340D);

        for (int index = 0; index < 4; index++) {
            HtmlRenderShape shape = FindColumnShape(rendered, "div#sizing-" + index);
            Assert.Equal(12D + columnAdvance * (index / blocksPerColumn), shape.X, 3);
            Assert.Equal(12D + 20D * (index % blocksPerColumn), shape.Y, 3);
            Assert.Equal(20D, shape.Height, 3);
        }
        HtmlRenderShape columns = FindColumnShape(rendered, "section#sizing-columns");
        Assert.Equal(expectedWidth, columns.Width, 3);
        Assert.Equal(expectedHeight, columns.Height, 3);
        Assert.Equal(expectedHeight, FindColumnShape(rendered, "div#sizing-after").Y, 3);
    }

    [Fact]
    public void HtmlColumns_IndefiniteSequentialFillContainsTheOrdinaryFloatTail() {
        const string html = "<style>body{margin:0;font:12px/20px Arial}p{margin:0}</style>"
            + "<section id='sizing-columns' style='width:220px;column-count:2;column-gap:20px;column-fill:auto;background:#ddd'>"
            + "<div id='sizing-float' style='float:left;width:40px;height:100px;background:red'></div>"
            + "<p>One</p><p>Two</p></section>"
            + "<div id='sizing-after' style='height:10px;background:black'></div>";

        HtmlRenderDocument rendered = RenderColumns(html, 240D);

        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        Assert.Equal(2, text.Length);
        Assert.Equal(40D, Assert.Single(text, item => item.Text == "One").X, 3);
        Assert.Equal(0D, Assert.Single(text, item => item.Text == "One").Y, 3);
        Assert.Equal(40D, Assert.Single(text, item => item.Text == "Two").X, 3);
        Assert.Equal(20D, Assert.Single(text, item => item.Text == "Two").Y, 3);
        Assert.Equal(100D, FindColumnShape(rendered, "div#sizing-float").Height, 3);
        Assert.Equal(100D, FindColumnShape(rendered, "section#sizing-columns").Height, 3);
        Assert.Equal(100D, FindColumnShape(rendered, "div#sizing-after").Y, 3);
    }

    private static string ColumnAutoSizingSource(string sizing, string thirdBlockStyle = "") =>
        "<style>body{margin:0}</style><section id='sizing-columns' "
        + "style='width:300px;column-count:2;column-gap:20px;background:#ddd;" + sizing + "'>"
        + string.Concat(Enumerable.Range(0, 4).Select(index =>
            "<div id='sizing-" + index + "' style='height:20px;background:blue;"
            + (index == 2 ? thirdBlockStyle : "") + "'></div>"))
        + "</section><div id='sizing-after' style='height:10px;background:black'></div>";
}
