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
}
