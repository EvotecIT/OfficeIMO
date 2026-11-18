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
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p>Icon <span style='display:inline-block;width:10px;height:20px;vertical-align:middle'></span> and "
            + "<span style='vertical-align:middle'>text</span> <span style='text-transform:full-width'>abc</span></p>"
            + "<table><tr><td style='vertical-align:middle'>cell</td></tr></table>"
            + "<div style='vertical-align:bottom'>Blocks ignore vertical-align.</div>",
            TextKeywordOptions);
        string[] details = rendered.NativeDiagnostics
            .Where(item => item.Code == HtmlRenderDiagnosticCodes.CssPropertyUnsupported)
            .Select(item => item.Detail ?? string.Empty)
            .ToArray();

        // Reported once per value; table cells apply vertical-align and blocks ignore it.
        Assert.Equal(new[] { "text-transform=full-width", "vertical-align=middle" }, details.OrderBy(item => item, StringComparer.Ordinal));
    }
}
