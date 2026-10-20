using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("block")]
    [InlineData("inline")]
    [InlineData("contents")]
    public void HtmlColumns_DescendantSpannerCrossesOrdinaryAncestorsWithoutLosingEitherContinuation(string display) {
        string before = display == "inline" ? "<span>B1</span><br><span>B2</span>" : "<p>B1</p><p>B2</p>";
        string after = display == "inline" ? "<span>C1</span><br><span>C2</span>" : "<p>C1</p><p>C2</p>";
        string tag = display == "inline" ? "span" : "div";
        HtmlRenderDocument rendered = RenderDescendantSpanners("<p>P1</p><p>P2</p>"
            + "<" + tag + " id='wrapper' style='display:" + display + ";color:#005599'>" + before
            + "<h2 id='spanner' class='spanner'>SPAN</h2>" + after + "</" + tag + "><p>T1</p><p>T2</p>");

        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "B1", "B2", "SPAN", "C1", "C2", "T1", "T2", "AFTER");
        HtmlRenderShape spanner = DescendantSpannerBackground(rendered, "h2#spanner");
        Assert.Equal(300D, spanner.Width, 3);
        Assert.Equal(0D, spanner.X, 3);
        Assert.Equal(40D, spanner.Y, 3);
        Assert.Equal(160D, DescendantSpannerText(rendered, "B1").X, 3);
        Assert.Equal(0D, DescendantSpannerText(rendered, "C1").X, 3);
        Assert.Equal(60D, DescendantSpannerText(rendered, "C1").Y, 3);
        Assert.Equal(120D, DescendantSpannerText(rendered, "T2").Y, 3);
        Assert.Equal(140D, DescendantSpannerText(rendered, "AFTER").Y, 3);
        Assert.Equal(OfficeColor.FromRgb(0, 0x55, 0x99), DescendantSpannerText(rendered, "SPAN").Color);
        Assert.Equal("heading-2", DescendantSpannerText(rendered, "SPAN").Source);
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_DescendantProjectionKeepsOriginalSelectorsInheritanceAndGeneratedSourceOnce() {
        HtmlRenderDocument rendered = RenderDescendantSpanners("<p>P1</p><p>P2</p><div id='outer'><p>B1</p>"
            + "<div id='inner'><p>B2</p><h2 id='spanner' class='spanner'>SPAN</h2><p>C1</p></div>"
            + "<p>C2</p></div><p>T1</p><p>T2</p>", extraCss: """
                #outer{font-weight:700;color:#005599}
                #inner{padding-left:10%;padding-right:10%}
                #inner>p:nth-child(3){font-weight:400;color:#880044}
                #outer::before{content:'G1';display:block;font-weight:400}
                #outer::after{content:'G2';display:block;font-weight:400}
                #inner::before{content:'G3';display:block;font-weight:400}
                #inner::after{content:'G4';display:block;font-weight:400}
                """);

        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "G1", "B1", "G3", "B2", "SPAN", "C1", "G4", "C2", "G2", "T1", "T2", "AFTER");
        Assert.Equal(300D, DescendantSpannerBackground(rendered, "h2#spanner").Width, 3);
        HtmlRenderText continuation = DescendantSpannerText(rendered, "C1");
        Assert.False(continuation.Font.IsBold);
        Assert.Equal(OfficeColor.FromRgb(0x88, 0, 0x44), continuation.Color);
        Assert.Equal(14D, continuation.X, 3);
        HtmlRenderText inherited = DescendantSpannerText(rendered, "B2");
        Assert.True(inherited.Font.IsBold);
        Assert.Equal(OfficeColor.FromRgb(0, 0x55, 0x99), inherited.Color);
        foreach ((string label, string source) in new[] {
                     ("G1", "div#outer::before"), ("G2", "div#outer::after"),
                     ("G3", "div#inner::before"), ("G4", "div#inner::after") }) {
            HtmlRenderText generated = DescendantSpannerText(rendered, label);
            Assert.Equal(source, generated.Source);
            Assert.False(generated.Font.IsBold);
        }
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(true, 36D, 56D, 138D)]
    [InlineData(false, 42D, 78D, 158D)]
    public void HtmlColumns_DescendantSpannerKeepsEmptyFirstAndLastSliceEdgesExactlyOnce(
        bool first, double spannerY, double tailY, double afterY) {
        string body = first
            ? "<p>P1</p><p>P2</p><div id='wrapper'><h2 id='spanner' class='spanner'>SPAN</h2>"
                + DescendantSpannerParagraphs("T", 4) + "</div>"
            : "<div id='wrapper'>" + DescendantSpannerParagraphs("P", 4)
                + "<h2 id='spanner' class='spanner'>SPAN</h2></div>" + DescendantSpannerParagraphs("T", 4);
        string edges = first ? "margin-top:6px;padding-top:8px" : "padding-bottom:8px;margin-bottom:6px";
        HtmlRenderDocument rendered = RenderDescendantSpanners(body,
            extraCss: "#wrapper{border:2px solid #008833;background:#d9f0d9;box-decoration-break:slice;" + edges + "}");

        AssertDescendantSpannerSourceOnce(rendered, Enumerable.Range(1, first ? 2 : 4).Select(index => "P" + index)
            .Concat(new[] { "SPAN" }).Concat(Enumerable.Range(1, 4).Select(index => "T" + index)).Append("AFTER").ToArray());
        HtmlRenderShape spanner = DescendantSpannerBackground(rendered, "h2#spanner");
        Assert.Equal(300D, spanner.Width, 3);
        Assert.Equal(spannerY, spanner.Y, 3);
        Assert.Equal(tailY, DescendantSpannerText(rendered, "T1").Y, 3);
        Assert.Equal(afterY, DescendantSpannerText(rendered, "AFTER").Y, 3);
        Assert.Contains(DescendantSpannerShapes(rendered), shape => shape.Source == "div#wrapper"
            && shape.Shape.FillColor == OfficeColor.FromRgb(0xd9, 0xf0, 0xd9));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_DescendantCloneRetainsSourceAndFullWidthWhileReportingUnqualifiedDecoration() {
        HtmlRenderDocument Render(string decoration) => RenderDescendantSpanners(
            "<div id='wrapper'>" + DescendantSpannerParagraphs("P", 4)
            + "<h2 id='spanner' class='spanner'>SPAN</h2>" + DescendantSpannerParagraphs("T", 6) + "</div>",
            extraCss: "#wrapper{margin-left:10%;width:80%;padding:8px 5%;border:2px solid #008833;"
                + "background:#d9f0d9;box-decoration-break:" + decoration + "}");
        HtmlRenderDocument sliced = Render("slice"), cloned = Render("clone");

        string[] labels = Enumerable.Range(1, 4).Select(index => "P" + index).Append("SPAN")
            .Concat(Enumerable.Range(1, 6).Select(index => "T" + index)).Append("AFTER").ToArray();
        AssertDescendantSpannerSourceOnce(sliced, labels);
        AssertDescendantSpannerSourceOnce(cloned, labels);
        Assert.Equal(300D, DescendantSpannerBackground(cloned, "h2#spanner").Width, 3);
        AssertColumnBreaksQualified(sliced);
        Assert.Contains(cloned.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported
            && diagnostic.Source == "div#wrapper" && diagnostic.Detail?.Contains("box-decoration-break=clone", StringComparison.Ordinal) == true);
        Assert.True(cloned.HasLoss);
        Assert.Throws<HtmlConversionException>(() => cloned.RequireNoLoss());
    }

    [Fact]
    public void HtmlColumns_DescendantAdjacentAndEmptySpannersDoNotCreateAnExtraColumnSet() {
        // The empty h2 retains the native fixture's authored .spanner background.
        HtmlRenderDocument rendered = RenderDescendantSpanners("<div id='wrapper'>"
            + DescendantSpannerParagraphs("P", 4) + "<h2 id='first' class='spanner'>S1</h2>"
            + "<h2 class='spanner'></h2><h2 id='last' class='spanner'>S2</h2>"
            + DescendantSpannerParagraphs("T", 2) + "</div>");

        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "P3", "P4", "S1", "S2", "T1", "T2", "AFTER");
        Assert.Equal(300D, DescendantSpannerBackground(rendered, "h2#first").Width, 3);
        Assert.Equal(300D, DescendantSpannerBackground(rendered, "h2#last").Width, 3);
        Assert.Equal(20D, DescendantSpannerText(rendered, "S2").Y - DescendantSpannerText(rendered, "S1").Y, 3);
        Assert.Equal(20D, DescendantSpannerText(rendered, "T1").Y - DescendantSpannerText(rendered, "S2").Y, 3);
        Assert.Equal(0D, DescendantSpannerText(rendered, "T2").X, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData("display:flow-root")]
    [InlineData("display:inline-block;width:140px;vertical-align:top")]
    [InlineData("transform:translateX(0px)")]
    public void HtmlColumns_DescendantExtractionStopsAtFormattingAndContainingBlockBarriers(string barrier) {
        HtmlRenderDocument rendered = RenderDescendantSpanners("<div id='barrier' style='" + barrier
            + "'><p>P1</p><h2 id='spanner' class='spanner'>SPAN</h2><p>T1</p></div>");

        AssertDescendantSpannerSourceOnce(rendered, "P1", "SPAN", "T1", "AFTER");
        HtmlRenderShape spanner = DescendantSpannerBackground(rendered, "h2#spanner");
        Assert.Equal(140D, spanner.Width, 3);
        Assert.Equal(20D, DescendantSpannerText(rendered, "SPAN").Y - DescendantSpannerText(rendered, "P1").Y, 3);
        Assert.Equal(20D, DescendantSpannerText(rendered, "T1").Y - DescendantSpannerText(rendered, "SPAN").Y, 3);
        Assert.Equal(60D, DescendantSpannerText(rendered, "AFTER").Y, 3);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlColumns_DescendantFittingKeepMovesWithoutCarryingItsEarlierAncestorSource(HtmlRenderIntentProfile profile) {
        HtmlRenderDocument rendered = RenderDescendantSpanners("<div id='wrapper'><p>P1</p><p>P2</p>"
            + "<div id='spanner' class='spanner' style='break-inside:avoid-page'>"
            + DescendantSpannerParagraphs("S", 3) + "</div>" + DescendantSpannerParagraphs("T", 8) + "</div>",
            profile, entry: "<p>E1</p><p>E2</p>");

        AssertDescendantSpannerSourceOnce(rendered, new[] { "E1", "E2", "P1", "P2", "S1", "S2", "S3" }
            .Concat(Enumerable.Range(1, 8).Select(index => "T" + index)).Append("AFTER").ToArray());
        Assert.Equal(1, DescendantSpannerTextPage(rendered, "P1").PageNumber);
        Assert.Equal(1, DescendantSpannerTextPage(rendered, "P2").PageNumber);
        Assert.All(new[] { "S1", "S2", "S3" }, label => Assert.Equal(2, DescendantSpannerTextPage(rendered, label).PageNumber));
        Assert.Equal(2, DescendantSpannerTextPage(rendered, "T4").PageNumber);
        Assert.Equal(3, DescendantSpannerTextPage(rendered, "T5").PageNumber);
        Assert.Equal(180D, DescendantSpannerText(rendered, "T3").X, 3);
        Assert.Equal(20D, DescendantSpannerText(rendered, "T5").X, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlColumns_DescendantNaturalSpannerWidowsDetermineTheTrailingPhysicalColumnBand(HtmlRenderIntentProfile profile) {
        string lines = string.Join("<br>", Enumerable.Range(1, 10).Select(index => "<span>S" + index + "</span>"));
        HtmlRenderDocument rendered = RenderDescendantSpanners("<div id='wrapper'><p>P1</p><p>P2</p>"
            + "<div id='spanner' class='spanner'><p style='orphans:2;widows:3'>" + lines + "</p></div>"
            + DescendantSpannerParagraphs("T", 18) + "</div>", profile, entry: "<p>E1</p><p>E2</p>");

        AssertDescendantSpannerSourceOnce(rendered, new[] { "E1", "E2", "P1", "P2" }
            .Concat(Enumerable.Range(1, 10).Select(index => "S" + index))
            .Concat(Enumerable.Range(1, 18).Select(index => "T" + index)).Append("AFTER").ToArray());
        Assert.Equal(new[] { 1, 1, 2, 2, 2, 2, 2, 3, 3, 3 }, Enumerable.Range(1, 10)
            .Select(index => DescendantSpannerTextPage(rendered, "S" + index).PageNumber).ToArray());
        Assert.All(Enumerable.Range(1, 10), index => Assert.Equal(20D, DescendantSpannerText(rendered, "S" + index).X, 3));
        Assert.Equal(3, DescendantSpannerTextPage(rendered, "T4").PageNumber);
        Assert.Equal(4, DescendantSpannerTextPage(rendered, "T5").PageNumber);
        Assert.Equal(4, DescendantSpannerTextPage(rendered, "T14").PageNumber);
        Assert.Equal(5, DescendantSpannerTextPage(rendered, "T15").PageNumber);
        Assert.Equal(5, DescendantSpannerTextPage(rendered, "AFTER").PageNumber);
        Assert.Equal(20D, DescendantSpannerText(rendered, "T1").X, 3);
        Assert.Equal(180D, DescendantSpannerText(rendered, "T3").X, 3);
        Assert.Equal(20D, DescendantSpannerText(rendered, "T1").Y - DescendantSpannerText(rendered, "S10").Y, 3);
        Assert.All(DescendantSpannerShapes(rendered).Where(shape => shape.Source == "div#spanner"),
            shape => Assert.Equal(300D, shape.Width, 3));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void HtmlColumns_DescendantAndDirectInternalPageForceOverridesAvoidAndRetainsTheTerminalBand(bool before, bool descendant) {
        string spanner = "<div id='spanner' class='spanner' style='break-inside:avoid-page'>"
            + "<p>S1</p><p style='" + (before ? string.Empty : "break-after:page") + "'>S2</p>"
            + "<p style='" + (before ? "break-before:page" : string.Empty) + "'>S3</p><p>S4</p></div>";
        string content = "<p>P1</p><p>P2</p>" + spanner + DescendantSpannerParagraphs("T", 12);
        foreach (HtmlRenderIntentProfile profile in new[] {
                     HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged,
                     HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderIntentProfile.ScreenSnapshotPaged }) {
            HtmlRenderDocument rendered = RenderDescendantSpanners(descendant ? "<div id='wrapper'>" + content + "</div>" : content, profile);

            AssertDescendantSpannerSourceOnce(rendered, new[] { "P1", "P2", "S1", "S2", "S3", "S4" }
                .Concat(Enumerable.Range(1, 12).Select(index => "T" + index)).Append("AFTER").ToArray());
            bool paged = profile is HtmlRenderIntentProfile.PrintPaged or HtmlRenderIntentProfile.ScreenMediaPaged;
            Assert.Equal(1, DescendantSpannerTextPage(rendered, "S1").PageNumber);
            Assert.Equal(paged ? 2 : 1, DescendantSpannerTextPage(rendered, "S3").PageNumber);
            Assert.Equal(paged ? 3 : 1, DescendantSpannerTextPage(rendered, "T7").PageNumber);
            if (paged) {
                Assert.Equal(20D, DescendantSpannerText(rendered, "S3").Y, 3);
                Assert.Equal(20D, DescendantSpannerText(rendered, "T1").Y - DescendantSpannerText(rendered, "S4").Y, 3);
                Assert.Equal(180D, DescendantSpannerText(rendered, "T4").X, 3);
            } else {
                Assert.Single(rendered.Pages);
                Assert.Equal(20D, DescendantSpannerText(rendered, "S3").Y - DescendantSpannerText(rendered, "S2").Y, 3);
            }
            AssertColumnBreaksQualified(rendered);
        }
    }

    private static HtmlRenderDocument RenderDescendantSpanners(string content,
        HtmlRenderIntentProfile profile = HtmlRenderIntentProfile.ScreenFullPage, string extraCss = "", string entry = "") {
        HtmlRenderOptions options = ColumnPageBandOptions();
        options.ViewportWidth = 800D;
        options.ViewportHeight = 1000D;
        options.DefaultFontSize = 14D;
        options.UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser;
        string html = ColumnBreakSource("<style>@page{size:700px 140px;margin:20px}"
            + "p,h2{margin:0;font:inherit;orphans:1;widows:1}"
            + "#columns{width:300px;column-count:2;column-gap:20px;column-fill:auto}"
            + ".spanner{column-span:all;margin:0;font-weight:700;background:#eee}" + extraCss + "</style>"
            + entry + "<section id='columns'>" + content + "</section><p id='after'>AFTER</p>");
        return HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, options)).Document;
    }

    private static string DescendantSpannerParagraphs(string prefix, int count) =>
        string.Concat(Enumerable.Range(1, count).Select(index => "<p>" + prefix + index + "</p>"));

    private static IEnumerable<HtmlRenderShape> DescendantSpannerShapes(HtmlRenderDocument rendered) =>
        rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderShape>();

    private static HtmlRenderShape DescendantSpannerBackground(HtmlRenderDocument rendered, string source) =>
        Assert.Single(DescendantSpannerShapes(rendered), shape => shape.Source == source
            && shape.Shape.FillColor == OfficeColor.FromRgb(0xee, 0xee, 0xee));

    private static HtmlRenderText DescendantSpannerText(HtmlRenderDocument rendered, string label) =>
        Assert.Single(rendered.Pages.SelectMany(page => DescendantSpannerOwnedText(page.Scene)), text => text.Text == label);

    private static HtmlRenderPage DescendantSpannerTextPage(HtmlRenderDocument rendered, string label) =>
        Assert.Single(rendered.Pages, page => DescendantSpannerOwnedText(page.Scene).Any(text => text.Text == label));

    private static void AssertDescendantSpannerSourceOnce(HtmlRenderDocument rendered, params string[] labels) {
        string[] scene = rendered.Pages.SelectMany(page => DescendantSpannerOwnedText(page.Scene))
            .SelectMany(text => text.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).ToArray();
        string[] logical = rendered.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(labels.OrderBy(value => value, StringComparer.Ordinal), scene.OrderBy(value => value, StringComparer.Ordinal));
        Assert.Equal(labels.OrderBy(value => value, StringComparer.Ordinal), logical.OrderBy(value => value, StringComparer.Ordinal));
    }

    private static IEnumerable<HtmlRenderText> DescendantSpannerOwnedText(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Artifact }
                or HtmlRenderLogicalTextGroup { Text.Length: 0 }) continue;
            if (visual is HtmlRenderText text) {
                if (text.Text.Length > 0) yield return text;
                continue;
            }
            IReadOnlyList<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup group => group.Visuals, HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals, HtmlRenderSemanticGroup group => group.Visuals,
                HtmlRenderLayoutRegion group => group.Visuals, HtmlRenderLogicalTextGroup group => group.Visuals,
                HtmlRenderFormField group => group.Visuals, _ => null
            };
            if (children == null) continue;
            foreach (HtmlRenderText child in DescendantSpannerOwnedText(children)) yield return child;
        }
    }
}
