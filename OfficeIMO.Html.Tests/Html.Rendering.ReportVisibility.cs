using OfficeIMO.Drawing;
using OfficeIMO.ContentSafety;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("<section hidden class='reveal'>REVEALED</section>", "block")]
    [InlineData("<p>Before <span hidden class='reveal'>REVEALED</span> after</p>", "inline")]
    [InlineData("<div style='display:flex'><div hidden class='reveal'>REVEALED</div></div>", "block")]
    [InlineData("<div style='display:grid;grid-template-columns:1fr 1fr'><div hidden class='reveal'>REVEALED</div><div>Other</div></div>", "block")]
    [InlineData("<table><tbody><tr hidden class='reveal'><td>REVEALED</td></tr></tbody></table>", "table-row")]
    [InlineData("<table><tbody><tr><td hidden class='reveal'>REVEALED</td></tr></tbody></table>", "table-cell")]
    public void HtmlReportVisibility_PrintDisplayOverridesHidden(string fragment, string display) {
        string html = "<style>@media print{.reveal{display:" + display + "!important}}</style>"
            + fragment + "<section hidden>PLAIN-HIDDEN</section><p style='display:none'>CSS-HIDDEN</p>";

        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            HonorCssPageRules = false, InteractiveFormControls = false
        });
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("REVEALED", text, StringComparison.Ordinal);
        Assert.DoesNotContain("PLAIN-HIDDEN", text, StringComparison.Ordinal);
        Assert.DoesNotContain("CSS-HIDDEN", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(HtmlRenderMode.Continuous, false)]
    [InlineData(HtmlRenderMode.Paged, true)]
    public void HtmlReportVisibility_RevealFollowsActiveMedia(HtmlRenderMode mode, bool revealed) {
        const string html = "<style>@media print{section[hidden]{display:block!important}}</style>"
            + "<section hidden>REVEALED</section><p>VISIBLE</p>";

        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = mode, Margins = HtmlRenderMargins.All(0D)
        });
        string text = string.Join(" ", document.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(item => item.Text));

        Assert.Contains("VISIBLE", text, StringComparison.Ordinal);
        Assert.Equal(revealed, text.Contains("REVEALED", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlReportVisibility_InlineDisplayOverridesHidden() {
        const string html = "<span hidden style='display:inline'>INLINE-REVEALED</span><span hidden>PLAIN-HIDDEN</span>";
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes()).ExtractText();

        Assert.Contains("INLINE-REVEALED", text, StringComparison.Ordinal);
        Assert.DoesNotContain("PLAIN-HIDDEN", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlReportVisibility_OutlineUsesTheSameVisibleHeadingText() {
        const string html = "<h1>VISIBLE-TITLE <span hidden>HIDDEN-TITLE</span>"
            + "<span hidden style='display:inline'>REVEALED-TITLE</span></h1>";
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes();
        string title = Assert.Single(OfficeIMO.Pdf.PdfInspector.Inspect(bytes).Outlines).Title;

        Assert.Contains("VISIBLE-TITLE", title, StringComparison.Ordinal);
        Assert.Contains("REVEALED-TITLE", title, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN-TITLE", title, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("hidden", false)]
    [InlineData("style='display:none'", false)]
    [InlineData("hidden style='display:inline'", true)]
    [InlineData("", true)]
    public void HtmlReportVisibility_LineBreakFollowsResolvedDisplay(string attributes, bool breaksLine) {
        HtmlRenderDocument document = HtmlRenderTestDriver.Render("<p>BEFORE<br " + attributes + ">AFTER</p>");
        HtmlRenderText[] text = document.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        HtmlRenderText before = Assert.Single(text, item => item.Text.Contains("BEFORE", StringComparison.Ordinal));
        HtmlRenderText after = Assert.Single(text, item => item.Text.Contains("AFTER", StringComparison.Ordinal));

        Assert.Equal(breaksLine, after.Y > before.Y);
    }

    [Theory]
    [InlineData("auto")]
    [InlineData("fixed")]
    public void HtmlReportVisibility_HiddenCellsDoNotParticipateInTracksOrSpans(string layout) {
        string start = "<table style='width:120px;table-layout:" + layout + "'><tr>";
        const string omitted = "<td hidden colspan='4' rowspan='2' style='width:1000px'>PLAIN-HIDDEN</td>"
            + "<th style='display:none;width:1000px'>CSS-HIDDEN</th>";
        const string end = "<td>FIRST</td></tr><tr><td>SECOND</td></tr></table>";
        var options = new HtmlRenderOptions { ViewportWidth = 140D, MaxTableColumns = 2 };
        HtmlRenderDocument actual = HtmlRenderTestDriver.Render(start + omitted + end, options);
        HtmlRenderDocument expected = HtmlRenderTestDriver.Render(start + end, options);
        HtmlRenderText[] actualText = actual.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        HtmlRenderText[] expectedText = expected.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();

        Assert.DoesNotContain(actualText, item => item.Text.Contains("HIDDEN", StringComparison.Ordinal));
        foreach (string value in new[] { "FIRST", "SECOND" }) {
            HtmlRenderText observed = Assert.Single(actualText, item => item.Text == value);
            HtmlRenderText reference = Assert.Single(expectedText, item => item.Text == value);
            Assert.Equal(reference.X, observed.X, 6);
            Assert.Equal(reference.Y, observed.Y, 6);
            Assert.Equal(reference.Width, observed.Width, 6);
        }
    }

    [Theory]
    [InlineData("div", "", "initial", "inline", true)]
    [InlineData("div", "", "unset", "inline", true)]
    [InlineData("div", "", "inherit", "block", true)]
    [InlineData("span", "", "inherit", "inline", true)]
    [InlineData("div", "display:flex", "inherit", "flex", true)]
    [InlineData("div", "", "revert", "none", false)]
    [InlineData("div", "", "revert-layer", "none", false)]
    public void HtmlReportVisibility_CssWideDisplayUsesInitialInheritedOrUserAgentValue(
        string parentTag, string parentStyle, string keyword, string expectedDisplay, bool revealed) {
        string html = "<" + parentTag + " style='" + parentStyle + "'><span id='target' hidden style='display:"
            + keyword + "'>REVEALED</span></" + parentTag + ">";
        var dom = HtmlDocumentParser.ParseDocument(html);
        HtmlComputedStyle computed = HtmlComputedStyleEngine.Compute(dom)[dom.QuerySelector("#target")!];
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes()).ExtractText();

        Assert.Equal(expectedDisplay, computed.GetValue("display"));
        Assert.Equal(revealed, text.Contains("REVEALED", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlReportVisibility_LayerRevertRetainsEarlierDisplayOverride() {
        const string html = "<style>@layer base,theme;@layer base{#target{display:block}}"
            + "@layer theme{#target{display:revert-layer}}</style><section id='target' hidden>REVEALED</section>";
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes()).ExtractText();

        Assert.Contains("REVEALED", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("initial")]
    [InlineData("unset")]
    [InlineData("inherit")]
    public void HtmlReportVisibility_SafetyCleanupPreservesRevealedText(string keyword) {
        string html = "<div><span hidden style='display:" + keyword + "'>VISIBLE-REVEALED</span>"
            + "<span hidden>PLAIN-HIDDEN</span><span hidden style='display:revert'>REVERT-HIDDEN</span></div>";
        OfficeContentSafetyReport report = HtmlContentSafety.Inspect(html);
        OfficeContentSafetyFinding[] concealed = report.Findings
            .Where(item => item.Kind == OfficeContentConcealmentKind.HiddenByProperty).ToArray();

        Assert.DoesNotContain(concealed, item => item.TextPreview.Contains("VISIBLE-REVEALED", StringComparison.Ordinal));
        Assert.Contains(concealed, item => item.TextPreview.Contains("PLAIN-HIDDEN", StringComparison.Ordinal));
        Assert.Contains(concealed, item => item.TextPreview.Contains("REVERT-HIDDEN", StringComparison.Ordinal));
        OfficeContentCleanupResult cleaned = HtmlContentSafety.RemoveSelected(html,
            new OfficeContentCleanupSelection(concealed.Select(item => item.Id)));
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(HtmlConversionDocument.Parse(
            System.Text.Encoding.UTF8.GetString(cleaned.Output)).ToPdfBytes()).ExtractText();

        Assert.Contains("VISIBLE-REVEALED", text, StringComparison.Ordinal);
        Assert.DoesNotContain("PLAIN-HIDDEN", text, StringComparison.Ordinal);
        Assert.DoesNotContain("REVERT-HIDDEN", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("block")]
    [InlineData("inherit")]
    public void HtmlReportVisibility_FidelityScorerCountsRevealedText(string display) {
        string source = "<div><span hidden style='display:" + display + "'>Visible content</span></div>";
        HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(source, "<div></div>");

        Assert.Equal(0D, score.Dimensions["text"]);
    }
}
