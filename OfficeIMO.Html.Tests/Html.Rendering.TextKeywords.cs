using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private static readonly HtmlRenderOptions TextKeywordOptions = new() {
        Mode = HtmlRenderMode.Continuous,
        ViewportWidth = 400D,
        Margins = HtmlRenderMargins.All(0D)
    };

    [Theory]
    [InlineData("<p><u style='text-decoration:none'>plain</u></p>", false, false)]
    [InlineData("<p><del style='text-decoration-line:none'>plain</del></p>", false, false)]
    [InlineData("<p><u style='text-decoration:line-through'>plain</u></p>", false, true)]
    [InlineData("<p><s style='text-decoration-color:red'>plain</s></p>", false, true)]
    public void HtmlRenderer_AuthoredTextDecorationReplacesTagDefaults(string html, bool underline, bool strikethrough) {
        HtmlRenderText text = Assert.Single(JustificationVisuals(HtmlRenderTestDriver.Render(html, TextKeywordOptions).Pages[0].Visuals)
            .OfType<HtmlRenderText>(), item => item.Text == "plain");

        Assert.Equal(underline, text.Font.IsUnderline);
        Assert.Equal(strikethrough, text.Font.IsStrikethrough);
    }

    [Theory]
    [InlineData("word-wrap:break-word", true)]
    // The legacy name is an alias: the later declaration wins either way round.
    [InlineData("overflow-wrap:normal;word-wrap:break-word", true)]
    [InlineData("word-wrap:break-word;overflow-wrap:normal", false)]
    // A reset is a declaration too, and restores normal instead of inheriting.
    [InlineData("word-wrap:break-word;overflow-wrap:initial", false)]
    [InlineData("overflow-wrap:break-word;word-wrap:unset", false)]
    public void HtmlRenderer_LegacyWordWrapCascadesAsOverflowWrap(string declarations, bool wraps) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:60px;" + declarations + "'>Supercalifragilisticexpialidocious</div>", TextKeywordOptions);
        HtmlRenderText[] fragments = JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>()
            .Where(text => !string.IsNullOrWhiteSpace(text.Text)).ToArray();

        Assert.Equal("Supercalifragilisticexpialidocious", string.Concat(fragments.Select(text => text.Text)));
        if (wraps) {
            Assert.True(fragments.Length > 1, "The word must wrap inside the 60px box.");
            Assert.All(fragments, text => Assert.True(text.X + (text.TextAdvanceWidth ?? text.Width) <= 60.5D));
        } else {
            Assert.Single(fragments);
        }
    }

    [Theory]
    // overflow-wrap is inherited: unset keeps the parent's value, initial restores normal.
    [InlineData("overflow-wrap:unset", true)]
    [InlineData("word-wrap:unset", true)]
    [InlineData("overflow-wrap:initial", false)]
    public void HtmlRenderer_OverflowWrapResetsFollowInheritance(string childDeclaration, bool wraps) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:60px;overflow-wrap:break-word'><p style='margin:0;" + childDeclaration + "'>Supercalifragilisticexpialidocious</p></div>",
            TextKeywordOptions);

        Assert.Equal(wraps, JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>()
            .Count(text => !string.IsNullOrWhiteSpace(text.Text)) > 1);
    }

    [Fact]
    public void HtmlRenderer_ForcedBreakKeepsPreservedSpacesOfAPreformattedInline() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='text-align-last:right;width:200px;margin:0'>A<span style='white-space:pre'>   </span><br>B</p>", TextKeywordOptions);

        Assert.True(JustifiedLines(rendered)[0].Right < 195D, "The preserved spaces stay at the end of the line.");
    }

    [Fact]
    public void HtmlRenderer_ReportsInterCharacterJustificationOnInlineText() {
        Assert.Contains("text-justify=inter-character", UnsupportedTextDetails(
            "<p style='text-align:justify;width:150px;margin:0'>alpha <span style='text-justify:inter-character'>one two three four five six seven eight nine</span></p>"));
    }

    [Theory]
    [InlineData("white-space:pre-line")]
    // Anonymous flex item text is laid out without DOM run collection.
    [InlineData("display:flex;flex-direction:column;white-space:pre-line")]
    public void HtmlRenderer_PreLineKeepsSegmentBreaksAndCollapsesSurroundingSpaces(string declarations) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='" + declarations + "'>first   line   \n   second line\r\nthird</div>", TextKeywordOptions);
        IReadOnlyList<JustifiedLine> lines = JustifiedLines(rendered);

        Assert.Equal(new[] { "first line", "second line", "third" }, lines.Select(line => line.Text).ToArray());
        Assert.All(lines, line => Assert.Equal(0D, line.Left, 1));
    }

    [Fact]
    public void HtmlRenderer_VendorCenterKeywordCentersText() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='text-align:-webkit-center;width:300px;margin:0'>Centered</p>", TextKeywordOptions);
        JustifiedLine line = Assert.Single(JustifiedLines(rendered));

        Assert.Equal(300D - line.Right, line.Left, 1);
        Assert.True(line.Left > 10D);
    }

    [Fact]
    public void HtmlRenderer_ReportsTextKeywordsThatLayoutDoesNotApply() {
        string[] details = UnsupportedTextDetails(
            "<p>Icon <span style='display:inline-block;width:10px;height:20px;vertical-align:middle'></span> and "
            + "<span style='vertical-align:middle'>text</span> <span style='text-transform:full-width'>abc</span></p>");

        // Reported once per value.
        Assert.Equal(new[] { "text-transform=full-width", "vertical-align=middle" }, details.OrderBy(item => item, StringComparer.Ordinal));
    }

    [Theory]
    // Table cells apply the keyword; blocks, floats, positioned boxes and flex or grid items ignore vertical-align.
    [InlineData("<p>Text <img style='float:right;vertical-align:middle;width:5px;height:5px' src='data:image/gif;base64,R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw=='></p>")]
    [InlineData("<p style='position:relative'>Text <span style='position:absolute;vertical-align:bottom'>abs</span></p>")]
    [InlineData("<style>div::before{content:'x';vertical-align:middle}</style><div style='display:flex'>item</div>")]
    // inter-character is reported only when a line is actually justified.
    [InlineData("<p style='text-align:justify;text-justify:inter-character'>Short</p>")]
    [InlineData("<table><tr><td style='vertical-align:middle'>cell</td></tr></table>")]
    [InlineData("<div style='vertical-align:bottom'>Blocks ignore vertical-align.</div>")]
    [InlineData("<div style='display:flex'><span style='vertical-align:middle'>Flex item</span></div>")]
    [InlineData("<div style='display:grid'><span style='vertical-align:text-top'>Grid item</span></div>")]
    public void HtmlRenderer_DoesNotReportKeywordsWhereNothingIsApproximated(string html) {
        Assert.Empty(UnsupportedTextDetails(html));
    }

    [Theory]
    [InlineData("<p style='text-align:justify;width:150px;margin:0'>漢字の文章は単語の間に空白がないので文字の間で両端揃えにする必要があります。</p>", "text-align=justify;cjk")]
    [InlineData("<p dir='rtl' style='text-align:justify;width:150px;margin:0'>שלום <span style='display:inline-block;width:5px;height:5px'></span> one two three four five six seven eight</p>", "text-align=justify;rtl-inline-boxes")]
    public void HtmlRenderer_ReportsJustifiedLinesLeftAtStart(string html, string detail) {
        Assert.Contains(detail, UnsupportedTextDetails(html));
    }

    private static string[] UnsupportedTextDetails(string html) =>
        HtmlRenderTestDriver.Render(html, TextKeywordOptions).NativeDiagnostics
            .Where(item => item.Code == HtmlRenderDiagnosticCodes.CssPropertyUnsupported)
            .Select(item => item.Detail ?? string.Empty)
            .ToArray();

    [Fact]
    public void HtmlRenderer_PreLineBreaksGeneratedContent() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>p::before{content:'A\\A B';white-space:pre-line}</style><p style='margin:0'></p>", TextKeywordOptions);

        Assert.Equal(new[] { "A", "B" }, JustifiedLines(rendered).Select(line => line.Text).ToArray());
    }

    [Fact]
    public void HtmlRenderer_PreLineDropsTheSpaceBeforeABreakInAnotherTextNode() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='white-space:pre-line;text-align-last:right;width:200px;margin:0'>A <b>\nB</b></p>", TextKeywordOptions);

        Assert.Equal(200D, JustifiedLines(rendered)[0].Right, 1);
        Assert.DoesNotContain(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>(), text => text.Text == "A ");
    }

    [Theory]
    [InlineData("<div>0000000000\n0000000000</div>")]
    [InlineData("<div class='generated'></div>")]
    public void HtmlRenderer_PreLineFlexItemsAreSizedByTheirLongestLine(string item) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>.generated::before{content:'0000000000\\A 0000000000'}</style>"
            + "<div style='display:flex;white-space:pre-line'>" + item
            + "<div style='width:10px;height:10px;background:#ff0000'></div></div>", TextKeywordOptions);
        JustifiedLine first = JustifiedLines(rendered)[0];
        HtmlRenderShape sibling = Assert.Single(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillColor is OfficeIMO.OfficeColor { R: 0xFF, G: 0, B: 0 });

        Assert.Equal(first.Right, sibling.X, 0);
    }

    [Theory]
    [InlineData("(text-justify: inter-word)", true)]
    [InlineData("(text-justify: inter-character)", false)]
    [InlineData("(text-justify: distribute)", false)]
    public void HtmlStyles_SupportsQueriesDoNotClaimInterCharacterJustification(string condition, bool expected) {
        Assert.Equal(expected, HtmlComputedStyleEngine.IsApplicableSupports(condition));
    }
}
