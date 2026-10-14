using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlZeroLineHeightTests {
    [Theory]
    [InlineData("mixed-size")]
    [InlineData("mixed-face")]
    [InlineData("logical-group")]
    [InlineData("inherited-length")]
    [InlineData("superscript")]
    [InlineData("subscript")]
    [InlineData("generated")]
    [InlineData("first-letter")]
    [InlineData("empty-atom")]
    [InlineData("image")]
    public void ZeroLineHeightRetainsGlyphsWithoutMovingFollowingContent(string scenario) {
        string html = Source(scenario);
        var document = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(document,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: Options())).Document;
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => Enumerate(page.Scene)).ToArray();
        HtmlRenderText[] text = visuals.OfType<HtmlRenderText>().ToArray();
        string expectedText = scenario is "superscript" or "subscript" ? "A2" : scenario is "empty-atom" or "image" ? "A" : "AB";

        Assert.Equal(expectedText, rendered.Text.Replace("\n", ""));
        Assert.Contains(text, run => run.LineHeight == 0D && run.Text.Length > 0);
        Assert.All(text, run => Assert.True(run.Height > 0D && !double.IsInfinity(run.Height) && !double.IsNaN(run.Height)));
        HtmlRenderShape following = Assert.Single(visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#following" && shape.Shape.FillColor.HasValue);
        Assert.Equal(20D, following.Y, 6);
        Assert.NotEmpty(rendered.Pages[0].CreateDrawing().Elements);
        rendered.RequireNoLoss();

        HtmlPdfRenderRequestResult result = document.RenderToPdfResult(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options: Options()));
        PdfReadDocument pdf = PdfReadDocument.Open(result.ToBytes());
        // Inspect the serialized spans: geometric reading order can place a
        // raised superscript before the adjoining baseline text.
        Assert.Equal(expectedText, string.Concat(pdf.Pages.SelectMany(page => page.GetTextSpans())
            .Select(span => span.Text)).Replace("\n", "").Replace(" ", ""));
        result.Output.RequireNoLoss();
        Assert.Equal(html, document.SourceHtml);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    public void ScreenPagingPreservesZeroLineHeightTextAndFollowingGeometry(HtmlRenderIntentProfile profile) {
        var document = HtmlConversionDocument.Parse(Source("empty-atom"));
        HtmlPdfRenderRequestResult result = document.RenderToPdfResult(HtmlRenderRequest.Create(
            profile, HtmlRenderEncoder.Pdf, options: Options()));
        HtmlRenderDocument rendered = result.RenderResult.Document;

        Assert.Equal("A", rendered.Text.Trim());
        Assert.Contains(rendered.Pages.SelectMany(page => Enumerate(page.Scene)).OfType<HtmlRenderText>(),
            text => text.LineHeight == 0D && text.Text == "A");
        Assert.Equal(20D, Assert.Single(rendered.Pages.SelectMany(page => Enumerate(page.Scene))
            .OfType<HtmlRenderShape>(), shape => shape.Source == "div#following" && shape.Shape.FillColor.HasValue).Y, 6);
        Assert.Equal("A", PdfReadDocument.Open(result.ToBytes()).ExtractText().Trim());
        result.Output.RequireNoLoss();
    }

    private static string Source(string scenario) {
        string content = scenario switch {
            "mixed-size" => "<div>A<span style='font-size:12px;line-height:0'>B</span></div>",
            "mixed-face" => "<div>A<span style='font:12px/0 Alternative'>B</span></div>",
            "logical-group" => "<div>A<span style='font-size:12px;line-height:0'>\u200eB</span></div>",
            "inherited-length" => "<div>A<span style='line-height:0px'><span style='font-size:12px'>B</span></span></div>",
            "superscript" => "<div>A<sup style='font-size:75%;line-height:0'>2</sup></div>",
            "subscript" => "<div>A<sub style='font-size:75%;line-height:0'>2</sub></div>",
            "generated" => "<style>#generated::before{content:'B';font-size:12px;line-height:0}</style><div>A<span id='generated'></span></div>",
            "first-letter" => "<style>p::first-letter{font-size:12px;line-height:0}</style><p>AB</p>",
            "empty-atom" => "<div style='line-height:0'>A<span style='display:inline-block;width:20px;height:20px;background:lime'></span></div>",
            "image" => "<div style='line-height:0'>A<img width='20' height='20' alt='' src='data:image/png;base64,"
                + "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='></div>",
            _ => throw new ArgumentOutOfRangeException(nameof(scenario))
        };
        return "<!doctype html><style>html,body{margin:0;padding:0;font:16px/20px Pinned}p{margin:0}</style>"
            + content + "<div id='following' style='width:10px;height:10px;background:blue'></div>";
    }

    private static HtmlRenderOptions Options() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 600D, ViewportHeight = 900D,
            PageSize = new OfficePageSize(600D / 96D, 900D / 96D),
            Margins = HtmlRenderMargins.All(0D), HonorCssPageRules = false,
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
            ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile()
        };
        byte[] font = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf"));
        options.Fonts.Add("Pinned", font);
        options.Fonts.Add("Alternative", font);
        return options;
    }

    private static IEnumerable<HtmlRenderVisual> Enumerate(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderSemanticGroup group => group.Visuals,
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                _ => null
            };
            if (children != null) foreach (HtmlRenderVisual child in Enumerate(children)) yield return child;
        }
    }
}
