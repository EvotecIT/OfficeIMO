using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("שלום ", "עולם", "םלוע םולש")]
    [InlineData("سلام ", "بكم", "\uFEE2\uFEDC\uFE91 \uFEE1\uFE8E\uFEE0\uFEB3")]
    public void HtmlRenderer_OrdersPlainRtlWordsAcrossStyledInlineRuns(
        string firstWord,
        string secondWord,
        string expectedVisualText) {
        string html = "<p style='margin:0;width:400px;direction:ltr'>"
            + "<span style='color:#ff0000'>" + firstWord + "</span>"
            + "<span style='color:#0000ff'>" + secondWord + "</span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderText[] text = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .OrderBy(static run => run.X)
            .ToArray();

        Assert.Equal(expectedVisualText, string.Concat(text.Select(static run => run.Text)));
        Assert.Contains(firstWord + secondWord, rendered.Text, StringComparison.Ordinal);
        Assert.All(text, static run => Assert.True(run.BidiVisualOrderResolved));
        Assert.Contains(text, static run => run.Color == OfficeColor.FromRgb(255, 0, 0));
        Assert.Contains(text, static run => run.Color == OfficeColor.FromRgb(0, 0, 255));
    }
}
