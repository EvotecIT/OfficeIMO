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

    [Theory]
    [InlineData("display:grid;grid-template-columns:1fr;align-items:stretch")]
    [InlineData("display:flex;flex-direction:column;align-items:stretch")]
    public void HtmlColumns_LongFirstItemRetainsItsContentHeightBeforeFollowingItems(string containerStyle) {
        string body = ColumnIntrinsicFlowBody(
            containerStyle, bCount: 12,
            leading: "<p>LEADING</p>");
        HtmlRenderDocument continuous = RenderContinuousColumnIntrinsicFlow(body);
        Assert.Equal(240D, FindColumnShape(continuous, "section#intrinsic-b").Height, 3);
        AssertColumnIntrinsicAfter(continuous, 1, 340D);

        HtmlRenderDocument paged = RenderColumnPageBands(body);

        AssertColumnIntrinsicFlowText(paged, 4, 12);
        AssertColumnPageBandTextOnce(paged, new[] { "LEADING" });
        Assert.Equal(3, ColumnPageBandTextPage(paged, "A-01").PageNumber);
        Assert.Equal(80D, Assert.Single(ColumnPageBandText(paged.Pages[2]), text => text.Text == "A-01").LayoutY, 3);
        Assert.Equal(80D, Assert.Single(ColumnPageBandText(paged.Pages[2]), text => text.Text == "A-03").LayoutY, 3);
        Assert.True(Assert.Single(ColumnPageBandText(paged.Pages[2]), text => text.Text == "A-03").X > 350D);
        AssertColumnIntrinsicAfter(paged, 4, 60D);
        AssertColumnPageBandNoFragmentLoss(paged);
    }

    [Fact]
    public void HtmlColumns_DefiniteColumnFlexHeightRetainsItsAutoBasisDuringProjection() {
        HtmlRenderDocument rendered = RenderColumnPageBands(ColumnIntrinsicFlowBody(
            "display:flex;flex-direction:column;height:200px;align-items:stretch"));

        AssertColumnIntrinsicFlowText(rendered, 4, 4);
        Assert.Equal(1, ColumnPageBandTextPage(rendered, "A-02").PageNumber);
        HtmlRenderText second = Assert.Single(ColumnPageBandText(rendered.Pages[0]), text => text.Text == "A-02");
        Assert.True(second.X > 350D);
        Assert.Equal(100D, second.LayoutY, 3);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "A-03").PageNumber);
        AssertColumnIntrinsicAfter(rendered, 3, 20D);
        AssertColumnPageBandNoFragmentLoss(rendered);
    }

    [Theory]
    [InlineData("start", 2)]
    [InlineData("stretch", 3)]
    public void HtmlColumns_RowFlexSeparatesProjectedCrossHeightFromStretchAllocation(string alignment, int afterPage) {
        string body = ColumnIntrinsicFlowBody("display:flex;align-items:" + alignment, aCount: 12,
            itemStyle: "flex:1 1 0;min-width:0");
        HtmlRenderDocument continuous = RenderContinuousColumnIntrinsicFlow(body);
        AssertColumnIntrinsicAfter(continuous, 1, 240D);

        HtmlRenderDocument paged = RenderColumnPageBands(body);

        AssertColumnIntrinsicFlowText(paged, 12, 4);
        AssertColumnIntrinsicAfter(paged, afterPage, 60D);
        AssertColumnPageBandNoFragmentLoss(paged);
    }

    [Theory]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, true)]
    [InlineData(false, true, false)]
    public void HtmlColumns_DeferredFirstContentCarriesItsOrdinaryWrapperInsets(bool grid, bool signedMargins, bool insets) {
        string html = "<style>.page-columns{height:auto}p{line-height:20px}</style>"
            + "<main style='" + (grid ? "display:grid;grid-template-columns:1fr" : "display:flex;flex-direction:column") + "'>"
            + "<div style='" + (insets ? "border:1px solid #335;padding:4px;" : "") + "order:2;"
            + (signedMargins ? "margin-top:-6px;margin-bottom:8px" : "margin:0")
            + "'><section class='page-columns'>" + ColumnIntrinsicParagraphs("A", 4) + "</section></div>"
            + "<div style='" + (insets ? "border:1px solid #335;padding:4px;" : "") + "order:1;"
            + (signedMargins ? "margin-top:4px;margin-bottom:6px" : "margin:0")
            + "'><section class='page-columns'>" + ColumnIntrinsicParagraphs("B", 4) + "</section></div></main><p>AFTER</p>";

        HtmlRenderDocument rendered = RenderColumnPageBands(html);

        AssertColumnIntrinsicFlowText(rendered, 4, 4);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "A-01").PageNumber);
        Assert.Equal(insets ? 25D : 20D, Assert.Single(ColumnPageBandText(rendered.Pages[1]), text => text.Text == "A-01").LayoutY, 3);
        Assert.Equal(insets ? 85D : 80D, Assert.Single(ColumnPageBandText(rendered.Pages[1]), text => text.Text == "A-04").LayoutY, 3);
        AssertColumnIntrinsicAfter(rendered, 3, 20D);
        AssertColumnPageBandNoFragmentLoss(rendered);
    }

    [Fact]
    public void HtmlColumns_WrapperBottomDecorationDoesNotMoveTheLastPaintedColumnLineAcrossItsAllocationTail() {
        string body = "<style>.page-columns{height:auto}p{line-height:20px}#intrinsic-a p{line-height:19px}</style>"
            + "<main style='display:flex;flex-direction:column;align-items:stretch'>"
            + "<div style='border:1px solid #335;padding:4px;order:2'>"
            + "<section id='intrinsic-a' class='page-columns'>" + ColumnIntrinsicParagraphs("A", 6) + "</section></div>"
            + "<div style='border:1px solid #335;padding:4px;order:1'>"
            + "<section class='page-columns'>" + ColumnIntrinsicParagraphs("B", 4) + "</section></div></main><p>AFTER</p>";

        HtmlRenderDocument rendered = RenderColumnPageBands(body);

        AssertColumnIntrinsicFlowText(rendered, 6, 4);
        Assert.Equal(101D, Assert.Single(ColumnPageBandText(rendered.Pages[1]), text => text.Text == "A-05").LayoutY, 3);
        Assert.Equal(25D, Assert.Single(ColumnPageBandText(rendered.Pages[1]), text => text.Text == "A-06").LayoutY, 3);
        AssertColumnIntrinsicAfter(rendered, 3, 44D);
        AssertColumnPageBandNoFragmentLoss(rendered);
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
