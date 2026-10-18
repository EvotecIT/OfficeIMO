using AngleSharp.Dom;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("transform:translate(30px,5px)", "transform", "translate(30px,5px)")]
    [InlineData("transform:scale(2)", "transform", "scale(2)")]
    [InlineData("scale:2", "scale", "2")]
    [InlineData("scale:.5;transform:rotateX(20deg)", "transform", "rotateX(20deg)")]
    public void HtmlTransform_NonReplacedInlinePreservesDeclarationsWithoutMovingPaintOrLinks(
        string declaration, string property, string expectedDeclaration) {
        string Source(string effect) => OrdinaryFontBoxSource(16D, 24D,
            "<span id='owner' style='background:lime;" + effect + "'>"
            + "<a href='https://example.test/inline'>Hgyp</a></span>TAIL");
        string html = Source(declaration);
        var parsed = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> computed = HtmlComputedStyleEngine.Compute(parsed);
        string preserved = computed[parsed.QuerySelector("#owner")!].GetValue(property);
        Assert.Equal(expectedDeclaration.ToLowerInvariant(), preserved.Replace(" ", string.Empty).ToLowerInvariant());
        HtmlRenderDocument expected = HtmlRenderTestDriver.Render(Source(string.Empty), OrdinaryBaselineOptions());
        HtmlRenderDocument actual = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());
        Assert.DoesNotContain(OrdinaryFontBoxVisuals(actual).OfType<HtmlRenderEffectGroup>(),
            group => group.Source == "span#owner");
        Assert.Equal(expected.Text, actual.Text);
        Assert.Equal(OrdinaryFontBoxShape(expected, "span#owner").X,
            OrdinaryFontBoxShape(actual, "span#owner").X, 6);
        Assert.Equal(OrdinaryFontBoxShape(expected, "span#owner").Width,
            OrdinaryFontBoxShape(actual, "span#owner").Width, 6);
        Assert.Equal(OrdinaryFontBoxShape(expected, "div#following").Y,
            OrdinaryFontBoxShape(actual, "div#following").Y, 6);
        byte[] before = TransformApplicabilityPdf(Source(string.Empty));
        byte[] after = TransformApplicabilityPdf(html);
        using var controlPdf = UglyToad.PdfPig.PdfDocument.Open(before);
        using var authoredPdf = UglyToad.PdfPig.PdfDocument.Open(after);
        var control = controlPdf.GetPage(1).Letters.OrderBy(letter => letter.StartBaseLine.X).ToArray();
        var authored = authoredPdf.GetPage(1).Letters.OrderBy(letter => letter.StartBaseLine.X).ToArray();
        Assert.Equal(control.Length, authored.Length);
        for (int index = 0; index < control.Length; index++) {
            Assert.Equal(control[index].Value, authored[index].Value);
            Assert.Equal(control[index].StartBaseLine.X, authored[index].StartBaseLine.X, 6);
            Assert.Equal(control[index].StartBaseLine.Y, authored[index].StartBaseLine.Y, 6);
        }
        AssertTransformLinkDelta(before, after, "https://example.test/inline", 0D, 0D);
        actual.RequireNoLoss();
    }

    [Theory]
    [InlineData("display:inline-block", "block")]
    [InlineData("float:left;width:80px", "block")]
    [InlineData("position:absolute;left:0;top:0;width:80px", "block")]
    [InlineData("", "flex")]
    [InlineData("", "grid")]
    [InlineData("", "contents-flex")]
    public void HtmlTransform_EligibleOriginalBoxRolesKeepTheirPaintTranslationAfterLayoutClones(
        string role, string parent) {
        string Source(bool transform) {
            string item = "<span id='owner' style='background:lime;" + role
                + (transform ? ";transform:translate(30px,5px)" : "") + "'>Hgyp</span>";
            if (parent == "contents-flex") item = "<div style='display:contents'>" + item + "</div>";
            return OrdinaryFontBoxSource(16D, 24D, "<div style='position:relative;display:"
                + (parent == "contents-flex" ? "flex" : parent)
                + ";width:160px;height:60px'>" + item + "</div>");
        }
        HtmlRenderDocument control = HtmlRenderTestDriver.Render(Source(false), OrdinaryBaselineOptions());
        HtmlRenderDocument actual = HtmlRenderTestDriver.Render(Source(true), OrdinaryBaselineOptions());
        Assert.Equal(OrdinaryFontBoxShape(control, "div#following").Y,
            OrdinaryFontBoxShape(actual, "div#following").Y, 6);
        using var controlPdf = UglyToad.PdfPig.PdfDocument.Open(TransformApplicabilityPdf(Source(false)));
        using var authoredPdf = UglyToad.PdfPig.PdfDocument.Open(TransformApplicabilityPdf(Source(true)));
        var expected = controlPdf.GetPage(1).Letters.ToArray();
        var translated = authoredPdf.GetPage(1).Letters.ToArray();
        Assert.Equal("Hgyp", string.Concat(translated.Select(letter => letter.Value)));
        Assert.Equal(expected.Length, translated.Length);
        for (int index = 0; index < expected.Length; index++) {
            // Each saved origin is rounded by the PDF writer to millipoints.
            Assert.InRange(Math.Abs((translated[index].StartBaseLine.X - expected[index].StartBaseLine.X) / .75D - 30D),
                0D, .001D / .75D + .000001D);
            Assert.InRange(Math.Abs((expected[index].StartBaseLine.Y - translated[index].StartBaseLine.Y) / .75D - 5D),
                0D, .001D / .75D + .000001D);
        }
        actual.RequireNoLoss();
    }

    [Theory]
    [InlineData("flex")]
    [InlineData("grid")]
    public void HtmlTransform_ParentlessAncestorResolutionKeepsBlockifiedFixedContainingBoxes(string display) {
        string html = "<style>body{margin:0;font:16px/24px Pinned}</style>"
            + "<div style='display:" + display + ";margin-left:40px;margin-top:20px;width:160px;height:60px'>"
            + "<span style='background:lime;transform:translate(30px,5px)'><span>"
            + "<span style='position:fixed;left:0;top:0;width:40px;height:24px'>FIX</span>"
            + "</span>Hgyp</span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(TransformApplicabilityPdf(html));
        var page = pdf.GetPage(1);
        var fixedFirst = Assert.Single(page.Letters, letter => letter.Value == "F");
        Assert.InRange(Math.Abs(fixedFirst.StartBaseLine.X / .75D - 70D),
            0D, .0005D / .75D + .000001D);
        Assert.InRange(Math.Abs((page.Height - fixedFirst.StartBaseLine.Y) / .75D - 42D),
            0D, .0005D / .75D + .000001D);
        Assert.Contains("Hgyp", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("FIX", rendered.Text, StringComparison.Ordinal);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlTransform_ClampedLegacyBoxRetainsOrdinaryInlineFixedViewportAnchoring() {
        string Source(bool transform) => "<style>body{margin:0;font:16px/24px Pinned}</style>"
            + "<div style='margin-left:40px;margin-top:20px;width:160px;display:-webkit-box;"
            + "-webkit-box-orient:vertical;-webkit-line-clamp:2'>"
            + "<span id='owner'" + (transform ? " style='transform:translate(30px,5px)'" : "")
            + "><span><span id='fixed' style='position:fixed;left:125px;top:4px;"
            + "width:40px;height:24px;background-color:red'>FIX</span></span>OWNER</span></div>";
        HtmlRenderDocument control = HtmlRenderTestDriver.Render(Source(false), OrdinaryBaselineOptions());
        HtmlRenderDocument actual = HtmlRenderTestDriver.Render(Source(true), OrdinaryBaselineOptions());
        HtmlRenderShape fixedBox = OrdinaryFontBoxShape(actual, "span#fixed");
        Assert.Equal(125D, fixedBox.X, 6);
        Assert.Equal(4D, fixedBox.Y, 6);
        Assert.Equal(OrdinaryFontBoxShape(control, "span#fixed").X, fixedBox.X, 6);
        Assert.Equal(OrdinaryFontBoxShape(control, "span#fixed").Y, fixedBox.Y, 6);
        Assert.DoesNotContain(OrdinaryFontBoxVisuals(actual).OfType<HtmlRenderEffectGroup>(),
            group => group.Source == "span#owner");
        Assert.Equal(control.Text, actual.Text);
        using var before = UglyToad.PdfPig.PdfDocument.Open(TransformApplicabilityPdf(Source(false)));
        using var after = UglyToad.PdfPig.PdfDocument.Open(TransformApplicabilityPdf(Source(true)));
        var expected = before.GetPage(1).Letters.ToArray();
        var observed = after.GetPage(1).Letters.ToArray();
        Assert.Equal(expected.Length, observed.Length);
        for (int index = 0; index < expected.Length; index++) {
            Assert.Equal(expected[index].Value, observed[index].Value);
            Assert.Equal(expected[index].StartBaseLine, observed[index].StartBaseLine);
        }
        actual.RequireNoLoss();
    }

    [Theory]
    [InlineData("float")]
    [InlineData("interrupted-block")]
    public void HtmlTransform_LaidOutRootAnchorsDoNotAddUntransformedPlacementSlotLinks(string role) {
        const string uri = "https://example.test/owned-anchor";
        string Source(bool transform) {
            string anchor = "<a id='owner' href='" + uri + "' style='"
                + (role == "float" ? "float:left" : "display:block")
                + ";width:90px;height:30px;background-color:lime"
                + (transform ? ";transform:translate(30px,5px)" : "") + "'>LINK</a>";
            return OrdinaryFontBoxSource(16D, 24D,
                role == "float" ? anchor : "<span>BEFORE" + anchor + "AFTER</span>");
        }
        HtmlRenderDocument control = HtmlRenderTestDriver.Render(Source(false), OrdinaryBaselineOptions());
        HtmlRenderDocument actual = HtmlRenderTestDriver.Render(Source(true), OrdinaryBaselineOptions());
        Assert.Equal(OrdinaryFontBoxShape(control, "div#following").Y,
            OrdinaryFontBoxShape(actual, "div#following").Y, 6);
        AssertTransformLinkDelta(TransformApplicabilityPdf(Source(false)),
            TransformApplicabilityPdf(Source(true)), uri, 30D, 5D);
        actual.RequireNoLoss();
    }

    [Fact]
    public void HtmlTransform_FloatedChildRetainsTheEnclosingOrdinaryAnchorFallback() {
        const string uri = "https://example.test/inherited-anchor";
        string html = OrdinaryFontBoxSource(16D, 24D,
            "<a href='" + uri + "'><span style='float:left;width:90px;height:30px'>LINK</span></a>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());
        PdfLinkAnnotation[] links = PdfReadDocument.Open(TransformApplicabilityPdf(html))
            .Pages.SelectMany(page => page.GetLinkAnnotations()).Where(link => link.Uri == uri).ToArray();
        Assert.NotEmpty(links);
        Assert.Contains("LINK", rendered.Text, StringComparison.Ordinal);
        rendered.RequireNoLoss();
    }

    private static byte[] TransformApplicabilityPdf(string html) => HtmlConversionDocument.Parse(html)
        .ToPdfBytes(new HtmlToPdfOptions(OrdinaryBaselineOptions()));

    private static void AssertTransformLinkDelta(byte[] before, byte[] after, string uri, double x, double y) {
        PdfLinkAnnotation[] Read(byte[] bytes) => PdfReadDocument.Open(bytes).Pages
            .SelectMany(page => page.GetLinkAnnotations()).Where(link => link.Uri == uri)
            .OrderBy(link => link.X1).ThenBy(link => link.Y1).ToArray();
        PdfLinkAnnotation[] control = Read(before), authored = Read(after);
        Assert.NotEmpty(authored);
        Assert.Equal(control.Length, authored.Length);
        for (int index = 0; index < control.Length; index++) {
            Assert.Equal(x, (authored[index].X1 - control[index].X1) / .75D, 6);
            Assert.Equal(x, (authored[index].X2 - control[index].X2) / .75D, 6);
            Assert.Equal(y, (control[index].Y1 - authored[index].Y1) / .75D, 6);
            Assert.Equal(y, (control[index].Y2 - authored[index].Y2) / .75D, 6);
        }
    }
}
