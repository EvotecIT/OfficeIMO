using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRendering_DefaultWordBreaking_DoesNotSplitAnUnbreakableWord() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:60px'>Availability</div>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 120D });

        HtmlRenderText[] words = rendered.Pages[0].Visuals
            .OfType<HtmlRenderText>()
            .Where(text => text.Text.Contains("Avail", StringComparison.Ordinal))
            .ToArray();

        HtmlRenderText word = Assert.Single(words);
        Assert.Equal("Availability", word.Text);
        Assert.True(word.TextAdvanceWidth > 60D);
    }

    [Theory]
    [InlineData("overflow-wrap:anywhere")]
    [InlineData("overflow-wrap:break-word")]
    [InlineData("word-break:break-all")]
    public void HtmlRendering_ExplicitEmergencyWordBreaking_SplitsAnOversizedWord(string declaration) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:60px;" + declaration + "'>Availability</div>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 120D });

        HtmlRenderText[] fragments = rendered.Pages[0].Visuals
            .OfType<HtmlRenderText>()
            .Where(text => text.Text.Length > 0)
            .ToArray();

        Assert.True(fragments.Length > 1);
        Assert.Equal("Availability", string.Concat(fragments.Select(fragment => fragment.Text)));
        Assert.True(fragments.Select(fragment => fragment.Y).Distinct().Count() > 1);
    }
}
