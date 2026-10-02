using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRendering_DefaultPatternsUseEachRunLanguageAndPreserveLogicalWords() {
        const string html = "<div lang='en-US' style='width:64px;font-size:16px;hyphens:auto'>"
            + "<p>representation</p><p lang='de-DE'>Silbentrennung</p>"
            + "<p lang='zz'>unsupportedlanguage unsupportedlanguage</p><p style='hyphens:none'>representation</p></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 200 });
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Contains(text, item => item.Text == "repre-");
        Assert.Contains(text, item => item.Text == "Silben-");
        Assert.Contains(text, item => item.Text == "unsupportedlanguage");
        Assert.Single(rendered.Diagnostics, item => item.Code == "HyphenationLanguageUnsupported");
        Assert.Contains(text, item => item.Text == "representation");
        Assert.Equal("representationSilbentrennungunsupportedlanguageunsupportedlanguagerepresentation",
            string.Concat(rendered.Text.Where(c => !char.IsWhiteSpace(c))));
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("zz")]
    public void HtmlRendering_CallerHyphenationOverridesEmbeddedPatternsIncludingNoBreaks(string language) {
        string html = "<div lang='" + language + "' style='width:64px;font-size:16px;hyphens:auto'>representation</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 200,
            TextHyphenationCallback = _ => Array.Empty<int>()
        });
        Assert.Equal("representation", Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>()).Text);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == "HyphenationLanguageUnsupported");
    }

    [Theory]
    [InlineData("lang=''")]
    [InlineData("xml:lang=''")]
    public void HtmlRendering_ExplicitUnknownLanguageClearsInheritedDictionary(string attribute) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div lang='en-US' style='width:64px;font-size:16px;hyphens:auto'>"
            + "<p " + attribute + "><span>representation</span></p></div>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 200 });
        Assert.Equal("representation", Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>()).Text);
    }
    [Theory]
    [InlineData("", "representation", "repre-", false)]
    [InlineData("lang='de-DE'", "Silbentrennung", "Silben-", false)]
    [InlineData("lang='zz'", "representation", "representation", false)]
    [InlineData("lang=''", "representation", "representation", false)]
    [InlineData("", "representation", "repre-", true)]
    [InlineData("lang='de-DE'", "Silbentrennung", "Silben-", true)]
    public void HtmlRendering_FirstLinePreservesOriginLanguageAndSelectedHyphen(
        string attribute, string word, string firstPaint, bool withFloat) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>p::first-line{color:red}</style>"
            + "<p lang='en-US' style='width:64px;font-size:16px;hyphens:auto'><span "
            + attribute + ">" + (withFloat ? "<i style='float:left;width:1px;height:1px'></i>" : "")
            + word + "</span></p>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 200 });
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal(firstPaint, text[0].Text);
        Assert.Equal(OfficeColor.Red, text[0].Color);
        Assert.All(text.Skip(1), item => Assert.Equal(OfficeColor.Black, item.Color));
        Assert.Equal(word, string.Concat(rendered.Text.Where(c => !char.IsWhiteSpace(c))));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlRendering_FirstLineRetainsWholeWordCallerAndManualBreaks(bool manual) {
        var observed = new List<string>();
        string word = manual ? "repre\u00ADsentation" : "representation";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>p::first-line{color:red}</style>"
            + "<p style='width:64px;font-size:16px;hyphens:auto'>" + word + "</p>",
            new HtmlRenderOptions {
                Mode = HtmlRenderMode.Continuous, ViewportWidth = 200,
                TextHyphenationCallback = token => {
                    observed.Add(token);
                    return token == "representation" ? new[] { 5, 10 } : Array.Empty<int>();
                }
            });
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal("repre-", text[0].Text);
        Assert.Contains(text, item => item.Text == "senta-" && item.Color == OfficeColor.Black);
        Assert.All(observed, token => Assert.Equal("representation", token));
        Assert.Equal("representation", string.Concat(rendered.Text.Where(c => !char.IsWhiteSpace(c))));
    }

    [Fact]
    public void HtmlRendering_FirstLinePrefersFittingManualHyphenOverLaterPatternBreak() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>p::first-line{color:red}</style>"
            + "<p lang='en-US' style='width:90px;font-size:16px;hyphens:auto'>repre\u00ADsentation</p>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 200 });
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal("repre-", text[0].Text);
        Assert.All(text.Skip(1), item => Assert.Equal(OfficeColor.Black, item.Color));
        Assert.Equal("representation", string.Concat(rendered.Text.Where(c => !char.IsWhiteSpace(c))));
    }

    [Theory]
    [InlineData("hyphenate-limit-zone:60px")]
    [InlineData("white-space:nowrap")]
    public void HtmlRendering_FirstLinePreservesWholeWordPolicies(string policy) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>div::first-line{color:red}</style>"
            + "<div lang='en-US' style='width:90px;font-size:12px;hyphens:auto;" + policy + "'>to show typography</div>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 200 });
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.DoesNotContain(text, item => item.Text.Contains('-'));
        Assert.Contains(text, item => item.Text.Contains("typography"));
        Assert.Equal("toshowtypography", string.Concat(rendered.Text.Where(c => !char.IsWhiteSpace(c))));
    }

}
