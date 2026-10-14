using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("hel<span></span>lo world")]
    [InlineData("hel<span style='text-transform:none' lang='tr'></span>lo world")]
    [InlineData("hel<span style='padding:1px;border:1px solid black'>lo</span> world")]
    [InlineData("hel<span id='target'></span>lo world")]
    [InlineData("hel<span style='unicode-bidi:embed;direction:ltr'>lo</span> world")]
    [InlineData("hel<span style='unicode-bidi:isolate;direction:ltr'>lo</span> world")]
    [InlineData("hel<span style='string-set:chapter content()'>lo</span> world")]
    [InlineData("hel<span style='string-set:chapter \"Y\"'></span>lo world")]
    [InlineData("hel<span style='position:running(chapter)'></span>lo world")]
    [InlineData("hel<span style='position:absolute'></span>lo world")]
    [InlineData("hel<span style='position:fixed'></span>lo world")]
    public void HtmlRender_CapitalizationIgnoresNonTextInlineMarkers(string content) {
        string html = "<p style='margin:0;text-transform:capitalize'>" + content + "</p>"
            + "<a href='#target'></a>";

        AssertInlineTransformText(html, "Hello World");
    }

    [Theory]
    [InlineData("Ο<span style='font-weight:bold'>Σ</span>", "ος")]
    [InlineData("ΟΣ<span style='font-weight:bold'>Α</span>", "οσα")]
    public void HtmlRender_ContextualLowercaseRetainsWordContextAcrossInlineMarkers(string content, string expected) {
        AssertInlineTransformText("<p lang='el' style='margin:0;text-transform:lowercase'>"
            + content + "</p>", expected);
    }

    [Theory]
    [InlineData("br")]
    [InlineData("atom")]
    [InlineData("image")]
    public void HtmlRender_ActualInlineContentBoundariesStillSeparateCapitalization(string boundary) {
        string separator = boundary switch {
            "br" => "<br>",
            "image" => "<img alt='' style='width:1px;height:1px' src='data:image/png;base64,"
                + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1)) + "'>",
            _ => "<span style='display:inline-block;width:1px;height:1px'></span>"
        };

        AssertInlineTransformText("<p style='margin:0;text-transform:capitalize'>hel"
            + separator + "lo</p>", "HelLo");
    }

    [Fact]
    public void HtmlRender_ActualTextTransformOverridesRetainTheirOwnContext() {
        AssertInlineTransformText("<p style='margin:0;text-transform:uppercase'>ab"
            + "<span style='text-transform:none'>cd</span>ef</p>", "ABcdEF");
    }

    [Fact]
    public void HtmlRender_ActualTextLanguageOverridesRetainTheirOwnCasing() {
        AssertInlineTransformText("<p lang='tr' style='margin:0;text-transform:uppercase'>i"
            + "<span lang='en'>i</span>i</p>", "İIİ");
    }

    [Theory]
    [InlineData("hel<span style='font-size:0'> </span>lo", "HelLo")]
    [InlineData("hel<span style='font-size:0'>hidden </span>lo", "HelLo")]
    [InlineData("hel<span style='font-size:0'></span>lo", "Hello")]
    [InlineData("<span style='font-size:0'><span style='font-size:16px'>hel</span> <span style='font-size:16px'>lo</span></span>", "HelLo")]
    public void HtmlRender_SuppressedTextPreservesCasingContextWithoutVisibleOrPdfText(string content, string expected) {
        AssertInlineTransformText("<p style='margin:0;text-transform:capitalize'>" + content + "</p>", expected);
    }

    [Fact]
    public void HtmlRender_SuppressedGeneratedTextRetainsItsBoundaryWithoutVisibleOrPdfText() {
        AssertInlineTransformText("<style>.zero::before{content:' ';font-size:0}</style>"
            + "<p style='margin:0;text-transform:capitalize'>hel<span class='zero'></span>lo</p>", "HelLo");
    }

    [Fact]
    public void HtmlRender_CasingOnlyGeneratedNewlineDoesNotConstrainInlineBoxSizing() {
        const string content = "<span id='box' style='display:inline-block;background:yellow'>"
            + "AAAA<span class='zero'></span>BBBB</span><span>after</span>";
        const string pseudo = "<style>.zero::before{font-size:0;white-space:pre;content:'\\a'}</style>";
        HtmlRenderOptions options = CreateInlineTransformOptions();
        HtmlRenderDocument Render(string markup) => HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(
            "<!doctype html><style>html,body{margin:0;padding:0;font:16px/20px Pinned}</style>" + markup),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: options)).Document;
        HtmlRenderVisual[] expected = EnumerateTextOverflowVisuals(Render(content).Pages[0].Scene).ToArray();
        HtmlRenderVisual[] actual = EnumerateTextOverflowVisuals(Render(pseudo + content).Pages[0].Scene).ToArray();
        HtmlRenderShape expectedBox = Assert.Single(expected.OfType<HtmlRenderShape>(), shape => shape.Source == "span#box");
        HtmlRenderShape actualBox = Assert.Single(actual.OfType<HtmlRenderShape>(), shape => shape.Source == "span#box");
        HtmlRenderText expectedFollowing = Assert.Single(expected.OfType<HtmlRenderText>(), text => text.Text == "after");
        HtmlRenderText actualFollowing = Assert.Single(actual.OfType<HtmlRenderText>(), text => text.Text == "after");

        Assert.Equal(expectedBox.Width, actualBox.Width, 6);
        Assert.Equal(expectedBox.Height, actualBox.Height, 6);
        Assert.Equal(expectedFollowing.X, actualFollowing.X, 6);
        Assert.Equal(expectedFollowing.Y, actualFollowing.Y, 6);
        AssertInlineTransformText(pseudo + content, "AAAABBBBafter");
    }

    [Fact]
    public void HtmlRender_SuppressedGeneratedBoxesDoNotPublishGhostText() {
        AssertInlineTransformText("<style>.block::before{content:'BlockGhost';font-size:0;display:block}"
            + ".atomic::before{content:'AtomicGhost';font-size:0;display:inline-block}</style>"
            + "<div class='block'></div><a class='atomic' href='https://example.test'></a><p>Visible</p>", "Visible");
    }

    private static void AssertInlineTransformText(string content, string expected) {
        string html = "<!doctype html><style>html,body{margin:0;padding:0;font:16px/20px Pinned}</style>" + content;
        HtmlRenderOptions renderOptions = CreateInlineTransformOptions();
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: renderOptions)).Document;
        string text = string.Concat(rendered.Pages.SelectMany(page => EnumerateTextOverflowVisuals(page.Scene))
            .OfType<HtmlRenderText>().OrderBy(visual => visual.Y).ThenBy(visual => visual.X)
            .Select(visual => visual.Text));

        Assert.Equal(expected, text);
        string pdfText = PdfCore.PdfReadDocument.Open(HtmlConversionDocument.Parse(html)
            .RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
                HtmlRenderEncoder.Pdf, options: renderOptions)).ToBytes()).ExtractText();
        // Physical inline edges and hard breaks can add extraction whitespace;
        // the contract here is casing and contextual letters across serializers.
        Assert.Equal(string.Concat(expected.Where(character => !char.IsWhiteSpace(character))),
            string.Concat(pdfText.Where(character => !char.IsWhiteSpace(character))));
    }

    private static HtmlRenderOptions CreateInlineTransformOptions() {
        var renderOptions = new HtmlRenderOptions {
            ViewportWidth = 600D, ViewportHeight = 900D, Margins = HtmlRenderMargins.All(0D),
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser
        };
        renderOptions.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
        return renderOptions;
    }
}
