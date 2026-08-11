using OfficeIMO.Html;
using OfficeIMO.Drawing;
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

    [Fact]
    public void HtmlRendering_EmergencyWrappingRunsWhenNoAuthoredHyphenationPointFits() {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 120D,
            TextHyphenationCallback = token => token == "typography" ? new[] { 8 } : Array.Empty<int>()
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:12px;font-size:12px;hyphens:auto;overflow-wrap:anywhere'>typography</div>",
            options);
        HtmlRenderText[] fragments = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();

        Assert.True(fragments.Select(fragment => fragment.Y).Distinct().Count() > 1);
        Assert.Equal("typography", string.Concat(fragments.Select(fragment => fragment.Text)));
        Assert.All(fragments, fragment => Assert.True(fragment.TextAdvanceWidth <= 12.001D));
    }

    [Fact]
    public void HtmlRendering_ManualSoftHyphensUseTheAuthorHyphenCharacterWithoutChangingLogicalText() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:42px;font-size:12px;hyphens:manual;hyphenate-character:\"·\"'>ty\u00ADpography</div>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 120D });

        HtmlRenderText[] fragments = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();

        Assert.True(fragments.Select(fragment => fragment.Y).Distinct().Count() > 1);
        Assert.Contains("·", string.Concat(fragments.Select(fragment => fragment.Text)), StringComparison.Ordinal);
        Assert.Equal("typography", string.Concat(rendered.Text.Where(character => !char.IsWhiteSpace(character))));
    }

    [Fact]
    public void HtmlRendering_AutomaticHyphenationUsesTheSharedLexiconAndCssCharacterLimits() {
        var lexicon = new OfficeTextHyphenationLexicon(new[] { "ty-pog-ra-phy" }, minimumPrefixLength: 1, minimumSuffixLength: 1);
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 120D }
            .UseTextHyphenationLexicon(lexicon);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:42px;font-size:12px;hyphens:auto;hyphenate-limit-chars:5 2 2'>typography</div>",
            options);
        HtmlRenderText[] fragments = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();

        Assert.True(fragments.Select(fragment => fragment.Y).Distinct().Count() > 1);
        Assert.Contains(fragments, fragment => fragment.Text.EndsWith("-", StringComparison.Ordinal));
        Assert.Equal("typography", string.Concat(rendered.Text.Where(character => !char.IsWhiteSpace(character))));
    }

    [Fact]
    public void HtmlRendering_HyphenateLimitLinesBoundsConsecutiveHyphenatedLineEnds() {
        var lexicon = new OfficeTextHyphenationLexicon(new[] { "ty-pog-ra-phy", "de-ter-min-is-tic" }, minimumPrefixLength: 1, minimumSuffixLength: 1);
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 120D }
            .UseTextHyphenationLexicon(lexicon);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:42px;font-size:12px;hyphens:auto;hyphenate-limit-lines:1'>typography deterministic</div>",
            options);
        string[] lines = rendered.Pages[0].Visuals.OfType<HtmlRenderText>()
            .GroupBy(fragment => fragment.Y)
            .OrderBy(group => group.Key)
            .Select(group => string.Concat(group.OrderBy(fragment => fragment.X).Select(fragment => fragment.Text)))
            .ToArray();

        Assert.True(lines.Length > 2);
        Assert.DoesNotContain(Enumerable.Range(1, lines.Length - 1), index =>
            lines[index - 1].EndsWith("-", StringComparison.Ordinal) && lines[index].EndsWith("-", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlRendering_HyphensNoneSuppressesSoftHyphensAndRemovesTheConditionalCharacter() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:42px;font-size:12px;hyphens:none'>ty\u00ADpography</div>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 120D });
        HtmlRenderText fragment = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>());

        Assert.Equal("typography", fragment.Text);
        Assert.Equal("typography", rendered.Text);
    }

    [Fact]
    public void HtmlRendering_EmptyHyphenateCharacterBreaksWithoutPaintingAnInsertedGlyph() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:42px;font-size:12px;hyphens:manual;hyphenate-character:\"\"'>ty\u00ADpography</div>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 120D });
        HtmlRenderText[] fragments = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();

        Assert.True(fragments.Select(fragment => fragment.Y).Distinct().Count() > 1);
        Assert.Equal("typography", string.Concat(fragments.Select(fragment => fragment.Text)));
    }

    [Fact]
    public void HtmlRendering_AutomaticHyphenationPrefersAnAuthorSoftHyphenThatFits() {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 140D,
            TextHyphenationCallback = token => token == "typography" ? new[] { 2 } : Array.Empty<int>()
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:50px;font-size:12px;hyphens:auto'>typo\u00ADgraphy</div>",
            options);
        HtmlRenderText firstLine = rendered.Pages[0].Visuals.OfType<HtmlRenderText>()
            .OrderBy(fragment => fragment.Y)
            .ThenBy(fragment => fragment.X)
            .First();

        Assert.Equal("typo-", firstLine.Text);
    }

    [Fact]
    public void HtmlRendering_HyphenateLimitLastAlwaysKeepsAWholeFinalWordTogetherWhenItFitsAFreshLine() {
        var lexicon = new OfficeTextHyphenationLexicon(new[] { "ty-pog-ra-phy" }, minimumPrefixLength: 1, minimumSuffixLength: 1);
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 180D }
            .UseTextHyphenationLexicon(lexicon);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:90px;font-size:12px;hyphens:auto;hyphenate-limit-last:always'>to show typography</div>",
            options);
        string[] lines = rendered.Pages[0].Visuals.OfType<HtmlRenderText>()
            .GroupBy(fragment => fragment.Y)
            .OrderBy(group => group.Key)
            .Select(group => string.Concat(group.OrderBy(fragment => fragment.X).Select(fragment => fragment.Text)))
            .ToArray();

        Assert.Equal(2, lines.Length);
        Assert.Equal("typography", lines[1]);
        Assert.DoesNotContain("-", lines[0], StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRenderOptions_ClonePreservesTheHyphenationCallback() {
        OfficeTextHyphenationCallback callback = token => token == "typography" ? new[] { 2 } : Array.Empty<int>();
        var options = new HtmlRenderOptions { TextHyphenationCallback = callback };

        HtmlRenderOptions clone = options.Clone();

        Assert.Same(callback, clone.TextHyphenationCallback);
    }
}
