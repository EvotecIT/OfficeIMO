using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_FirstLetterRestoresOnlyItsOriginalZeroFontFragment() {
        const string content = "<style>p::first-letter{font-size:32px}</style>"
            + "<p style='margin:0;font-size:0;text-transform:capitalize'>Hidden "
            + "<span style='font-size:16px'>visible</span></p>";
        string html = FirstLetterSource(content);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(document,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: CreateInlineTransformOptions())).Document;
        HtmlRenderText[] text = FirstLetterText(rendered);

        Assert.Equal("HVisible", string.Concat(text.Select(item => item.Text)));
        Assert.Equal(32D, Assert.Single(text, item => item.Text == "H").Font.Size, 3);
        Assert.All(text.Where(item => item.Text != "H"), item => Assert.Equal(16D, item.Font.Size, 3));
        Assert.Equal(html, document.SourceHtml);
        rendered.RequireNoLoss();
        AssertInlineTransformText(content, "HVisible");
    }

    [Theory]
    [InlineData("32px", "HVisible", 32D)]
    [InlineData("2rem", "HVisible", 32D)]
    [InlineData("200%", "Visible", 0D)]
    [InlineData("2em", "Visible", 0D)]
    [InlineData("0", "Visible", 0D)]
    public void HtmlRender_FirstLetterUsesItsEffectiveFontWithoutSelectingLaterText(string size, string expected, double restoredSize) {
        string content = "<style>p::first-letter{font-size:" + size + "}</style>"
            + "<p style='margin:0;font-size:0;text-transform:capitalize'>Hidden "
            + "<span style='font-size:16px'>visible</span></p>";
        HtmlRenderText[] text = FirstLetterText(RenderFirstLetter(content));

        Assert.Equal(expected, string.Concat(text.Select(item => item.Text)));
        if (restoredSize > 0D) Assert.Equal(restoredSize, Assert.Single(text, item => item.Text == "H").Font.Size, 3);
        Assert.All(text.Where(item => item.Text != "H"), item => Assert.Equal(16D, item.Font.Size, 3));
    }

    [Fact]
    public void HtmlRender_FirstLetterInheritsTheRestoredDescendantsFontAndColor() {
        const string content = "<style>p::first-letter{font-size:200%}</style>"
            + "<p style='margin:0;font-size:0;color:red'>"
            + "<span style='font-size:16px;color:blue'>Hello</span></p>";
        HtmlRenderText[] text = FirstLetterText(RenderFirstLetter(content));

        Assert.Equal(32D, Assert.Single(text, item => item.Text == "H").Font.Size, 3);
        Assert.All(text, item => Assert.Equal(OfficeColor.Blue, item.Color));
        Assert.Equal(16D, Assert.Single(text, item => item.Text == "ello").Font.Size, 3);
        AssertInlineTransformText(content, "Hello");
    }

    [Fact]
    public void HtmlRender_FirstLetterWithoutSpecifiedFontRetainsTheDescendantsFont() {
        const string content = "<style>p::first-letter{color:blue}</style>"
            + "<p style='margin:0;font-size:0'><span style='font-size:16px'>Hello</span></p>";
        HtmlRenderText[] text = FirstLetterText(RenderFirstLetter(content));

        Assert.Equal(16D, Assert.Single(text, item => item.Text == "H").Font.Size, 3);
        Assert.Equal(OfficeColor.Blue, Assert.Single(text, item => item.Text == "H").Color);
        Assert.Equal("Hello", string.Concat(text.Select(item => item.Text)));
    }

    [Fact]
    public void HtmlRender_FirstLetterRestoresZeroFontGeneratedBeforeText() {
        const string content = "<style>p::before{content:'Hidden ';font-size:0}"
            + "p::first-letter{font-size:32px}</style>"
            + "<p style='margin:0;text-transform:capitalize'>visible</p>";
        HtmlRenderText[] text = FirstLetterText(RenderFirstLetter(content));

        Assert.Equal("HVisible", string.Concat(text.Select(item => item.Text)));
        Assert.Equal(32D, Assert.Single(text, item => item.Text == "H").Font.Size, 3);
        Assert.All(text.Where(item => item.Text != "H"), item => Assert.Equal(16D, item.Font.Size, 3));
        AssertInlineTransformText(content, "HVisible");
    }

    [Theory]
    [InlineData("display:inline-block", "Hidden ", "Visible")]
    [InlineData("display:inline-block", "Hidden", "visible")]
    [InlineData("display:inline-block;width:min-content", "Hidden ", "Visible")]
    [InlineData("display:inline-block;width:min-content", "Hidden\u00a0", "Visible")]
    [InlineData("display:inline-block;width:max-content", "Hidden ", "Visible")]
    [InlineData("display:inline-block;width:fit-content", "Hidden ", "Visible")]
    public void HtmlRender_FirstLetterIntrinsicSizingMatchesTheRestoredPaintAndFollowingText(string sizing, string hidden, string visible) {
        string wrapper = "<style>#box{margin:0;font-size:0;background:yellow;" + sizing + "}</style>";
        string actual = wrapper + "<style>#box::first-letter{font-size:32px}</style>"
            + "<p id='box' style='text-transform:capitalize'>" + hidden
            + "<span style='font-size:16px'>visible</span></p><span>after</span>";
        // The suppressed source space permits wrapping without advancing. At
        // min-content width, an explicit break supplies the independent control.
        string expectedBreak = sizing.Contains("width:min-content") && hidden.EndsWith(" ", StringComparison.Ordinal)
            ? "<br>"
            : string.Empty;
        string expected = wrapper + "<p id='box'><span style='font-size:32px'>H</span>"
            + expectedBreak + "<span style='font-size:16px'>" + visible + "</span></p><span>after</span>";
        HtmlRenderVisual[] actualVisuals = EnumerateTextOverflowVisuals(RenderFirstLetter(actual).Pages[0].Scene).ToArray();
        HtmlRenderVisual[] expectedVisuals = EnumerateTextOverflowVisuals(RenderFirstLetter(expected).Pages[0].Scene).ToArray();
        HtmlRenderShape actualBox = Assert.Single(actualVisuals.OfType<HtmlRenderShape>(), item => item.Source == "p#box");
        HtmlRenderShape expectedBox = Assert.Single(expectedVisuals.OfType<HtmlRenderShape>(), item => item.Source == "p#box");
        HtmlRenderText actualFollowing = Assert.Single(actualVisuals.OfType<HtmlRenderText>(), item => item.Text == "after");
        HtmlRenderText expectedFollowing = Assert.Single(expectedVisuals.OfType<HtmlRenderText>(), item => item.Text == "after");

        Assert.Equal(expectedBox.Width, actualBox.Width, 6);
        Assert.Equal(expectedBox.Height, actualBox.Height, 6);
        Assert.Equal(expectedFollowing.X, actualFollowing.X, 6);
        Assert.Equal(expectedFollowing.Y, actualFollowing.Y, 6);
    }

    [Theory]
    [InlineData("<br>")]
    [InlineData("<span style='display:inline-block;width:1px;height:1px'></span>")]
    public void HtmlRender_FirstLetterDoesNotSelectTextAfterTheFirstLineWasBlocked(string prefix) {
        string content = "<p id='box' style='margin:0;display:inline-block;background:yellow'>" + prefix + "Hello</p>";
        HtmlRenderDocument rendered = RenderFirstLetter("<style>p::first-letter{font-size:32px}</style>" + content);
        HtmlRenderDocument control = RenderFirstLetter(content);
        HtmlRenderText[] text = FirstLetterText(rendered);
        HtmlRenderShape actualBox = Assert.Single(EnumerateTextOverflowVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderShape>(), item => item.Source == "p#box");
        HtmlRenderShape controlBox = Assert.Single(EnumerateTextOverflowVisuals(control.Pages[0].Scene)
            .OfType<HtmlRenderShape>(), item => item.Source == "p#box");

        Assert.Equal("Hello", string.Concat(text.Select(item => item.Text)));
        Assert.All(text, item => Assert.Equal(16D, item.Font.Size, 3));
        Assert.Equal(controlBox.Width, actualBox.Width, 6);
        Assert.Equal(controlBox.Height, actualBox.Height, 6);
    }

    [Fact]
    public void HtmlRender_NonbreakingSpaceRetainsUnbrokenMinContentAndPaint() {
        const string wrapper = "<style>#box{margin:0;display:inline-block;width:min-content;background:yellow}</style>";
        const string content = "<p id='box'>A\u00a0B</p><span>after</span>";
        HtmlRenderDocument rendered = RenderFirstLetter(wrapper + content);
        HtmlRenderDocument control = RenderFirstLetter(wrapper
            + "<p id='box' style='white-space:nowrap'>A\u00a0B</p><span>after</span>");
        HtmlRenderDocument ordinarySpace = RenderFirstLetter(wrapper
            + "<p id='box'>A B</p><span>after</span>");
        HtmlRenderVisual[] actualVisuals = EnumerateTextOverflowVisuals(rendered.Pages[0].Scene).ToArray();
        HtmlRenderVisual[] controlVisuals = EnumerateTextOverflowVisuals(control.Pages[0].Scene).ToArray();
        HtmlRenderShape actualBox = Assert.Single(actualVisuals.OfType<HtmlRenderShape>(), item => item.Source == "p#box");
        HtmlRenderShape controlBox = Assert.Single(controlVisuals.OfType<HtmlRenderShape>(), item => item.Source == "p#box");
        HtmlRenderShape spaceBox = Assert.Single(EnumerateTextOverflowVisuals(ordinarySpace.Pages[0].Scene)
            .OfType<HtmlRenderShape>(), item => item.Source == "p#box");
        HtmlRenderText actualFollowing = Assert.Single(actualVisuals.OfType<HtmlRenderText>(), item => item.Text == "after");
        HtmlRenderText controlFollowing = Assert.Single(controlVisuals.OfType<HtmlRenderText>(), item => item.Text == "after");

        Assert.Equal(controlBox.Width, actualBox.Width, 6);
        Assert.Equal(controlBox.Height, actualBox.Height, 6);
        Assert.Equal(controlFollowing.X, actualFollowing.X, 6);
        Assert.Equal(controlFollowing.Y, actualFollowing.Y, 6);
        Assert.True(actualBox.Width > spaceBox.Width);
        Assert.True(actualBox.Height < spaceBox.Height);
        Assert.Contains(FirstLetterText(rendered), item => item.Text == "A\u00a0B");
        rendered.RequireNoLoss();
        AssertInlineTransformText(wrapper + content, "A\u00a0Bafter");
    }

    private static string FirstLetterSource(string content) =>
        "<!doctype html><style>html,body{margin:0;padding:0;font:16px/20px Pinned}</style>" + content;

    private static HtmlRenderDocument RenderFirstLetter(string content) =>
        HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(FirstLetterSource(content)),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: CreateInlineTransformOptions())).Document;

    private static HtmlRenderText[] FirstLetterText(HtmlRenderDocument rendered) =>
        rendered.Pages.SelectMany(page => EnumerateTextOverflowVisuals(page.Scene)).OfType<HtmlRenderText>().ToArray();
}
