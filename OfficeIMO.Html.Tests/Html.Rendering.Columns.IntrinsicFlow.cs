using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(4, 2, 80D)]
    [InlineData(12, 4, 40D)]
    [InlineData(24, 6, 80D)]
    public void HtmlColumns_PagedColumnFlexRetainsIntrinsicItemSpaceAfterColumnProjection(
        int aCount, int afterPage, double afterY) {
        HtmlRenderDocument rendered = RenderColumnPageBands(ColumnIntrinsicFlowBody(
            "display:flex;flex-direction:column;align-items:stretch", aCount));

        AssertColumnIntrinsicFlowText(rendered, aCount, 4);
        AssertColumnIntrinsicAfter(rendered, afterPage, afterY);
        AssertColumnPageBandNoFragmentLoss(rendered);
    }

    [Theory]
    [InlineData("start")]
    [InlineData("stretch")]
    public void HtmlColumns_PagedGridAutoRowsRetainIntrinsicItemSpaceAfterColumnProjection(string alignment) {
        HtmlRenderDocument rendered = RenderColumnPageBands(ColumnIntrinsicFlowBody(
            "display:grid;grid-template-columns:1fr;grid-auto-rows:auto;align-items:" + alignment));

        AssertColumnIntrinsicFlowText(rendered, 4, 4);
        AssertColumnIntrinsicAfter(rendered, 2, 80D);
        AssertColumnPageBandNoFragmentLoss(rendered);
    }

    [Fact]
    public void HtmlColumns_GridLongFirstRowRetainsItsContentHeightBeforeFollowingRows() {
        string body = ColumnIntrinsicFlowBody(
            "display:grid;grid-template-columns:1fr;align-items:stretch", bCount: 12,
            leading: "<p>LEADING</p>");
        HtmlRenderDocument continuous = RenderContinuousColumnIntrinsicFlow(body);
        Assert.Equal(240D, FindColumnShape(continuous, "section#intrinsic-b").Height, 3);
        AssertColumnIntrinsicAfter(continuous, 1, 340D);

        HtmlRenderDocument paged = RenderColumnPageBands(body);

        AssertColumnIntrinsicFlowText(paged, 4, 12);
        AssertColumnPageBandTextOnce(paged, new[] { "LEADING" });
        AssertColumnIntrinsicAfter(paged, 4, 60D);
        AssertColumnPageBandNoFragmentLoss(paged);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlColumns_ColumnFlexZeroBasisRespectsTheAutomaticContentMinimum(bool zeroMinimum) {
        HtmlRenderDocument rendered = RenderContinuousColumnIntrinsicFlow(ColumnIntrinsicFlowBody(
            "display:flex;flex-direction:column;align-items:stretch",
            itemStyle: "flex:1 1 0px;" + (zeroMinimum ? "min-height:0" : "")));

        AssertColumnIntrinsicFlowText(rendered, 4, 4);
        HtmlRenderText after = Assert.Single(ColumnPageBandText(Assert.Single(rendered.Pages)), text => text.Text == "AFTER");
        if (zeroMinimum) Assert.InRange(after.LayoutY, 0D, 0.05D);
        else Assert.Equal(160D, after.LayoutY, 3);
    }

    [Theory]
    [InlineData(false, 80D)]
    [InlineData(true, 60D)]
    public void HtmlColumns_DefiniteColumnFlexHeightDoesNotReplaceTheAutomaticItemMinimum(
        bool zeroMinimum, double expectedItemHeight) {
        string body = "<style>.page-columns{height:auto}p{line-height:20px}</style>"
            + "<main id='intrinsic-container' style='display:flex;flex-direction:column;height:60px;background:#ddd'>"
            + "<section id='intrinsic-a' class='page-columns' style='flex:0 1 80px;background:blue;"
            + (zeroMinimum ? "min-height:0" : "") + "'>" + ColumnIntrinsicParagraphs("A", 4)
            + "</section></main><p>AFTER</p>";

        HtmlRenderDocument rendered = RenderContinuousColumnIntrinsicFlow(body);

        Assert.Equal(60D, FindColumnShape(rendered, "main#intrinsic-container").Height, 3);
        Assert.Equal(expectedItemHeight, FindColumnShape(rendered, "section#intrinsic-a").Height, 3);
        AssertColumnPageBandTextOnce(rendered, Enumerable.Range(1, 4).Select(index => "A-" + index.ToString("D2")).Append("AFTER"));
        AssertColumnIntrinsicAfter(rendered, 1, 60D);
    }

    [Fact]
    public void HtmlColumns_StretchedFixedGridRowSuppliesThePercentageChildHeight() {
        string body = "<style>.page-columns{height:auto}p{line-height:20px}</style>"
            + "<main id='intrinsic-container' style='display:grid;grid-template-columns:1fr;grid-auto-rows:80px;align-items:stretch;background:#ddd'>"
            + "<section class='page-columns'><div style='height:50%'>" + ColumnIntrinsicParagraphs("C", 3)
            + "</div><p>TAIL</p></section></main><p>AFTER</p>";

        HtmlRenderDocument rendered = RenderContinuousColumnIntrinsicFlow(body);

        AssertColumnPageBandTextOnce(rendered, new[] { "C-01", "C-02", "C-03", "TAIL", "AFTER" });
        Assert.Equal(40D, Assert.Single(ColumnPageBandText(Assert.Single(rendered.Pages)), text => text.Text == "TAIL").LayoutY, 3);
        Assert.Equal(80D, FindColumnShape(rendered, "main#intrinsic-container").Height, 3);
        AssertColumnIntrinsicAfter(rendered, 1, 80D);
    }

    private static string ColumnIntrinsicFlowBody(string containerStyle, int aCount = 4, int bCount = 4,
        string itemStyle = "", string leading = "") =>
        "<style>.page-columns{height:auto}p{line-height:20px;orphans:1;widows:1}</style>" + leading
        + "<main style='" + containerStyle + "'>"
        + "<section id='intrinsic-a' class='page-columns' style='order:2;background:#ddd;" + itemStyle + "'>"
        + ColumnIntrinsicParagraphs("A", aCount) + "</section>"
        + "<section id='intrinsic-b' class='page-columns' style='order:1;background:#eee;" + itemStyle + "'>"
        + ColumnIntrinsicParagraphs("B", bCount) + "</section></main><p>AFTER</p>";

    private static string ColumnIntrinsicParagraphs(string prefix, int count) =>
        string.Concat(Enumerable.Range(1, count).Select(index => "<p>" + prefix + "-" + index.ToString("D2") + "</p>"));

    private static HtmlRenderDocument RenderContinuousColumnIntrinsicFlow(string body) {
        HtmlRenderOptions options = ColumnPageBandOptions();
        options.Mode = HtmlRenderMode.Continuous;
        return HtmlRenderEngine.Render(HtmlConversionDocument.Parse(ColumnPageBandSource(body)), options);
    }

    private static void AssertColumnIntrinsicFlowText(HtmlRenderDocument rendered, int aCount, int bCount) =>
        AssertColumnPageBandTextOnce(rendered, Enumerable.Range(1, aCount).Select(index => "A-" + index.ToString("D2"))
            .Concat(Enumerable.Range(1, bCount).Select(index => "B-" + index.ToString("D2"))).Append("AFTER"));

    private static void AssertColumnIntrinsicAfter(HtmlRenderDocument rendered, int pageNumber, double y) {
        HtmlRenderPage page = ColumnPageBandTextPage(rendered, "AFTER");
        Assert.Equal(pageNumber, page.PageNumber);
        Assert.Equal(y, Assert.Single(ColumnPageBandText(page), text => text.Text == "AFTER").LayoutY, 3);
    }
}
