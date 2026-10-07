using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlTextIndentTests {
    [Theory]
    [InlineData("24px", 24D)]
    [InlineData("-24px", -24D)]
    [InlineData("10%", 20D)]
    [InlineData("calc(10% + 4px)", 24D)]
    public void IndentMovesOnlyFirstLine(string indent, double expected) {
        var text = Render("<div style='width:200px;text-indent:" + indent + "'>First<br>Second</div>");
        Assert.Equal(expected, Find(text, "First").X, 3);
        Assert.Equal(0D, Find(text, "Second").X, 3);
    }

    [Fact]
    public void InheritedLengthRetainsParentFontButPercentageUsesChildWidth() {
        var length = Render("<div style='font-size:20px;text-indent:2em'><div style='font-size:10px;width:100px'>Child</div></div>");
        var percent = Render("<div style='width:200px;text-indent:10%'><div style='width:100px'>Child</div></div>");
        Assert.Equal(40D, Find(length, "Child").X, 3);
        Assert.Equal(10D, Find(percent, "Child").X, 3);
    }

    [Theory]
    [InlineData("0")]
    [InlineData("0.0")]
    [InlineData("-0")]
    [InlineData("initial")]
    public void ChildCanResetInheritedIndent(string reset) {
        var text = Render("<div style='text-indent:24px'><div style='text-indent:" + reset + "'>Child</div></div>");
        Assert.Equal(0D, Find(text, "Child").X, 3);
    }

    [Fact]
    public void PositiveIndentReducesFirstLineWrappingWidth() {
        const string content = "alpha beta gamma delta epsilon zeta eta theta";
        var normal = Render("<div style='width:100px'>" + content + "</div>");
        var indented = Render("<div style='width:100px;text-indent:40px'>" + content + "</div>");
        Assert.Equal(40D, indented[0].X, 3);
        Assert.True(FirstLine(indented).Length < FirstLine(normal).Length);
        Assert.All(indented, item => Assert.True(item.X + item.TextAdvanceWidth <= 100.01D));
        Assert.Equal(0D, indented.First(item => item.Y > indented[0].Y).X, 3);
    }

    [Theory]
    [InlineData("24px", 176D)]
    [InlineData("-24px", 224D)]
    [InlineData("240px", -40D)]
    public void RtlIndentConsumesSpaceAtRightEdge(string indent, double firstRight) {
        var text = Render("<div style='direction:rtl;width:200px;text-indent:" + indent + "'>First<br>Second</div>");
        var first = Find(text, "First");
        var second = Find(text, "Second");
        Assert.Equal(firstRight, first.X + first.TextAdvanceWidth!.Value, 3);
        Assert.Equal(200D, second.X + second.TextAdvanceWidth!.Value, 3);
    }

    [Fact]
    public void FloatAndIndentBothContributeToFirstLineStart() {
        var text = Render("<div style='width:200px;text-indent:24px'><a id='start'></a><span style='float:left;width:30px;height:40px'></span>First<br>Second</div>");
        Assert.Equal(54D, Find(text, "First").X, 3);
        Assert.Equal(30D, Find(text, "Second").X, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnonymousTextAfterBlockDoesNotRestartParentIndent(bool inlineWrapper) {
        const string content = "First<div style='text-indent:0'>Middle</div>Last";
        var text = Render("<div style='text-indent:24px'>" + (inlineWrapper ? "<span>" + content + "</span>" : content) + "</div>");
        Assert.Equal(24D, Find(text, "First").X, 3);
        Assert.Equal(0D, Find(text, "Middle").X, 3);
        Assert.Equal(0D, Find(text, "Last").X, 3);
    }

    [Fact]
    public void LaterAnonymousBlockPreservesInlineBlockInheritance() {
        var text = Render("<div style='text-indent:24px'><div style='text-indent:0'>First</div><span style='display:inline-block;width:100px'>Child</span>Last</div>");
        Assert.Equal(24D, Find(text, "Child").X, 3);
    }

    [Fact]
    public void BibliographyHangingIndentAlignsContinuationWithPadding() {
        var text = Render("<ol style='list-style:none;margin:0;padding:0'><li style='font-size:16px;padding-left:1.5em;text-indent:-1.5em'><div>Author<br>Title</div></li></ol>");
        Assert.Equal(0D, Find(text, "Author").X, 3);
        Assert.Equal(24D, Find(text, "Title").X, 3);
    }

    [Theory]
    [InlineData("overflow-wrap:anywhere", "abcdefghijklmnopqrstuvwx")]
    [InlineData("word-break:break-all", "abcdefghijklmnopqrstuvwx")]
    [InlineData("hyphens:manual", "abc\u00addef\u00adghi\u00adjkl\u00admno\u00adpqr\u00adstu\u00advwx")]
    public void TokenBreakingUsesIndentedFirstLineAndFullFollowingLines(string css, string content) {
        var text = Render("<div style='width:100px;text-indent:40px;" + css + "'>" + content + "</div>");
        Assert.Equal(40D, text[0].X, 3);
        Assert.True(text.Select(item => item.Y).Distinct().Count() > 1);
        Assert.All(text, item => Assert.True(item.X + item.TextAdvanceWidth <= 100.01D));
        Assert.Equal(0D, text.First(item => item.Y > text[0].Y).X, 3);
        Assert.Equal(content.Replace("\u00ad", ""), string.Concat(text.Select(item => item.Text)).Replace("-", ""));
    }

    [Fact]
    public void PageContinuationDoesNotIndentAgainWhenPageWidthChanges() {
        string words = string.Join(" ", Enumerable.Range(0, 40).Select(index => "word" + index.ToString("D2")));
        var document = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            "<style>@page{size:3in 2in;margin:0.25in}@page:first{size:2in 2in;margin:0.5in}p{margin:0;font-size:16px;line-height:20px;text-indent:24px}</style><p>" + words + "</p>"),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.True(document.Pages.Count > 1);
        Assert.Equal(72D, document.Pages[0].Visuals.OfType<HtmlRenderText>().First().X, 3);
        Assert.Equal(24D, document.Pages[1].Visuals.OfType<HtmlRenderText>().First().X, 3);
        string actual = string.Join(" ", document.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(item => item.Text));
        Assert.Equal(words, string.Join(" ", actual.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)));
    }

    [Theory]
    [InlineData("24px hanging")]
    [InlineData("24px each-line")]
    public void UnsupportedIndentModesAreReported(string value) {
        var document = HtmlRenderEngine.Render(HtmlConversionDocument.Parse("<div style='text-indent:" + value + "'>Text</div>"));
        Assert.Contains(document.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TextIndentValueUnsupported);
    }

    [Theory]
    [InlineData("24px", true)]
    [InlineData("-10%", true)]
    [InlineData("24px hanging", false)]
    [InlineData("24px each-line", false)]
    public void SupportsMatchesImplementedIndentModes(string value, bool supported) {
        Assert.Equal(supported, HtmlComputedStyleEngine.IsApplicableSupports("(text-indent:" + value + ")"));
    }

    private static string FirstLine(HtmlRenderText[] text) => string.Concat(text.Where(item => item.Y == text[0].Y).Select(item => item.Text));
    private static HtmlRenderText Find(HtmlRenderText[] text, string value) => Assert.Single(text, item => item.Text == value);
    private static HtmlRenderText[] Render(string html) => HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html),
        new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D) })
        .Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
}
