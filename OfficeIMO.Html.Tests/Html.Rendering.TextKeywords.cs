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

    [Fact]
    public void HtmlRenderer_LegacyWordWrapBreaksLongWordsLikeOverflowWrap() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:60px;word-wrap:break-word'>Supercalifragilisticexpialidocious</div>", TextKeywordOptions);
        HtmlRenderText[] fragments = JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>()
            .Where(text => !string.IsNullOrWhiteSpace(text.Text)).ToArray();

        Assert.True(fragments.Length > 1, "The word must wrap inside the 60px box.");
        Assert.All(fragments, text => Assert.True(text.X + (text.TextAdvanceWidth ?? text.Width) <= 60.5D));
        Assert.Equal("Supercalifragilisticexpialidocious", string.Concat(fragments.Select(text => text.Text)));
    }

    [Fact]
    public void HtmlRenderer_PreLineKeepsSegmentBreaksAndCollapsesSurroundingSpaces() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='white-space:pre-line'>first   line   \n   second line\r\nthird</div>", TextKeywordOptions);
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
            + "<table><tr><td style='vertical-align:middle'>cell</td></tr></table>",
            TextKeywordOptions);
        string[] details = rendered.NativeDiagnostics
            .Where(item => item.Code == HtmlRenderDiagnosticCodes.CssPropertyUnsupported)
            .Select(item => item.Detail ?? string.Empty)
            .ToArray();

        // Reported once per value; table cells apply vertical-align and are not reported.
        Assert.Equal(new[] { "text-transform=full-width", "vertical-align=middle" }, details.OrderBy(item => item, StringComparer.Ordinal));
    }
}
