using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("block", HtmlRenderMode.Continuous)]
    [InlineData("block", HtmlRenderMode.Paged)]
    [InlineData("inline", HtmlRenderMode.Continuous)]
    [InlineData("contents", HtmlRenderMode.Paged)]
    [InlineData("flex", HtmlRenderMode.Paged)]
    [InlineData("grid", HtmlRenderMode.Paged)]
    public void HtmlDisclosure_ClosedBodyDoesNotParticipateInLayout(string display, HtmlRenderMode mode) {
        string html = "<details style='display:" + display + "'><summary>CLOSED-SUMMARY</summary>HIDDEN-DIRECT"
            + "<p style='display:block!important'>HIDDEN-BLOCK</p><summary>HIDDEN-SECOND-SUMMARY</summary>"
            + "<span style='position:absolute;left:0;top:0'>HIDDEN-POSITIONED</span></details>"
            + "<details open style='display:" + display + "'><summary>OPEN-SUMMARY</summary><p>OPEN-BODY</p></details>"
            + "<details open><summary>OUTER-SUMMARY</summary><details><summary>INNER-SUMMARY</summary>HIDDEN-INNER</details></details>"
            + "<details><summary>NESTED-CLOSED-SUMMARY</summary><details open><summary>HIDDEN-NESTED-SUMMARY</summary>HIDDEN-NESTED-BODY</details></details>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = mode });
        string text = string.Join(" ", rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(item => item.Text));

        foreach (string visible in new[] { "CLOSED-SUMMARY", "OPEN-SUMMARY", "OPEN-BODY", "OUTER-SUMMARY", "INNER-SUMMARY", "NESTED-CLOSED-SUMMARY" }) {
            Assert.Contains(visible, text, StringComparison.Ordinal);
        }
        Assert.DoesNotContain("HIDDEN-", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlDisclosure_PdfRetainsPreparedPrintBodyAndOmitsClosedBodyDestinations() {
        const string html = "<style>@media print{.actions{display:none!important}}</style>"
            + "<button class='actions'>HIDDEN-ACTION</button><details><summary>VISIBLE-SUMMARY</summary>HIDDEN-DIRECT"
            + "<h2 id='closed-target'>HIDDEN-HEADING</h2><a href='https://evotec.xyz/closed'>HIDDEN-LINK</a></details>"
            + "<details open><summary>PRINT-SUMMARY</summary><p>PRINT-BODY</p><h2 id='print-target'>PRINT-HEADING</h2></details>";
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { InteractiveFormControls = false });
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(bytes).ExtractText();
        var inspection = OfficeIMO.Pdf.PdfInspector.Inspect(bytes);

        Assert.Contains("VISIBLE-SUMMARY", text, StringComparison.Ordinal);
        Assert.Contains("PRINT-BODY", text, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN-", text, StringComparison.Ordinal);
        Assert.Equal("PRINT-HEADING", Assert.Single(inspection.Outlines).Title);
    }

    [Theory]
    [InlineData("inline-flex")]
    [InlineData("inline-grid")]
    public void HtmlDisclosure_ClosedBodyDoesNotSizeIntrinsicTracks(string display) {
        string start = "<div style='display:" + display + ";grid-template-columns:auto auto'><details><summary>SUMMARY</summary>";
        const string hidden = "HIDDEN-UNBREAKABLE-TEXT-THAT-MUST-NOT-SIZE-THE-TRACK<div style='width:1000px'>HIDDEN-BLOCK</div>";
        const string end = "</details><div>SIBLING</div></div>";
        var options = new HtmlRenderOptions { ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) };
        HtmlRenderText[] actual = HtmlRenderTestDriver.Render(start + hidden + end, options).Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        HtmlRenderText[] expected = HtmlRenderTestDriver.Render(start + end, options).Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();

        foreach (string value in new[] { "SUMMARY", "SIBLING" }) {
            HtmlRenderText observed = Assert.Single(actual, item => item.Text == value);
            HtmlRenderText reference = Assert.Single(expected, item => item.Text == value);
            Assert.Equal(reference.X, observed.X, 6);
            Assert.Equal(reference.Y, observed.Y, 6);
            Assert.Equal(reference.Width, observed.Width, 6);
        }
    }

    [Fact]
    public void HtmlDisclosure_ClosedTextDoesNotSizeTableColumns() {
        const string start = "<table style='width:400px;table-layout:auto'><tr><td><details><summary>SUMMARY</summary>";
        const string end = "</details></td><td>SIBLING</td></tr></table>";
        AssertDisclosureGeometry(start + "HIDDEN-UNBREAKABLE-TEXT-THAT-MUST-NOT-SIZE-THE-COLUMN" + end,
            start + end, "SUMMARY", "SIBLING");
    }

    [Fact]
    public void HtmlDisclosure_ClosedTableImageDoesNotConsumeResourceBudget() {
        const string hidden = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII=";
        const string visible = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg==";
        string html = "<table style='width:100px'><tr><td><details><summary>S</summary><img src='data:image/png;base64,"
            + hidden + "'></details></td><td><img src='data:image/png;base64," + visible + "'></td></tr></table>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 120D, ViewportHeight = 40D, Margins = HtmlRenderMargins.All(0D), MaxResourceCount = 1
        });

        Assert.Equal(Convert.FromBase64String(visible), Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>()).Bytes);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded);
    }

    [Theory]
    [InlineData("position:absolute;right:0;top:0", false)]
    [InlineData("position:absolute;right:0;top:0", true)]
    [InlineData("position:fixed;right:0;top:0", false)]
    [InlineData("float:left", false)]
    [InlineData("float:right", false)]
    [InlineData("display:inline-block", false)]
    public void HtmlDisclosure_ClosedTextDoesNotSizeShrinkToFitBoxes(string style, bool directDisclosure) {
        string start = "<div style='position:relative;width:400px;height:80px'>"
            + (directDisclosure ? "<details style='" + style + "'>" : "<div style='" + style + "'><details>")
            + "<summary>SUMMARY</summary>";
        string end = "</details>" + (directDisclosure ? "" : "</div>") + "AFTER</div>";
        AssertDisclosureGeometry(start + "HIDDEN-UNBREAKABLE-TEXT-THAT-MUST-NOT-SIZE-THE-BOX" + end,
            start + end, "SUMMARY", "AFTER");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlDisclosure_ClosedFootnoteDoesNotAdvanceVisibleNumbering(bool authoredCounters) {
        string css = "@page{size:240px 160px;margin:10px}body,p{margin:0;font-size:12px;line-height:16px}.note{float:footnote}";
        if (authoredCounters) css += "body{counter-reset:footnote}.note::footnote-call{content:'[' counter(footnote) ']'}"
            + ".note::footnote-marker{content:counter(footnote,upper-roman) '.'}";
        string html = "<style>" + css + "</style><details><summary>SUMMARY</summary>"
            + "<span class='note'>HIDDEN-NOTE</span></details><p>Body<span id='visible-note' class='note'>VISIBLE-NOTE</span></p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true
        });
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();

        Assert.Equal(authoredCounters ? "[1]" : "1", Assert.Single(text, item => item.Source == "span#visible-note:footnote-call").Text);
        Assert.Equal(authoredCounters ? "I." : "1", Assert.Single(text, item => item.Source == "span#visible-note:footnote-marker").Text);
        Assert.Contains(text, item => item.Text == "VISIBLE-NOTE");
        Assert.DoesNotContain(text, item => item.Text.Contains("HIDDEN-NOTE", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlDisclosure_ClosedBodyDoesNotChangeVisibleGeneratedCounterState() {
        const string html = "<style>body{counter-reset:section}.step{counter-increment:section}"
            + ".step::before{content:'STEP-' counter(section)}</style>"
            + "<details><summary>SUMMARY</summary><div class='step'>HIDDEN-STEP</div></details>"
            + "<div class='step'>VISIBLE-STEP</div>";
        HtmlRenderText[] text = HtmlRenderTestDriver.Render(html).Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();

        string visibleText = string.Concat(text.Select(item => item.Text));
        Assert.Contains("STEP-1", visibleText, StringComparison.Ordinal);
        Assert.DoesNotContain("STEP-2", visibleText, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN-STEP", visibleText, StringComparison.Ordinal);
    }

    private static void AssertDisclosureGeometry(string actualHtml, string expectedHtml, params string[] values) {
        var options = new HtmlRenderOptions { ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) };
        HtmlRenderText[] actual = HtmlRenderTestDriver.Render(actualHtml, options).Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        HtmlRenderText[] expected = HtmlRenderTestDriver.Render(expectedHtml, options).Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        foreach (string value in values) {
            HtmlRenderText observed = Assert.Single(actual, item => item.Text == value);
            HtmlRenderText reference = Assert.Single(expected, item => item.Text == value);
            Assert.Equal(reference.X, observed.X, 6);
            Assert.Equal(reference.Y, observed.Y, 6);
            Assert.Equal(reference.Width, observed.Width, 6);
        }
    }
}
