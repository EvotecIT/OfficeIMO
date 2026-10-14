using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("table")]
    [InlineData("tr")]
    [InlineData("td")]
    public void HtmlTableGeometry_AuthoredHeightIsAMinimum(string heightOwner) {
        string Height(string tag) => tag == heightOwner ? "height:100px;" : string.Empty;
        string html = TableGeometrySource("<table id='sized' style='" + Height("table") + "width:200px;margin:0;border-spacing:0'>"
            + "<tr style='" + Height("tr") + "'><td id='cell' style='" + Height("td")
            + "padding:0;background:lime'>Sized</td></tr></table><div id='next' style='background:red;height:10px'>Next</div>");

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderShape cell = TableGeometryShape(rendered, "td#cell");
        HtmlRenderShape next = TableGeometryShape(rendered, "div#next");

        Assert.Equal(100D, cell.Height, 3);
        Assert.Equal(100D, next.Y, 3);
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData("table", "content-box", 130D, 100D)]
    [InlineData("table", "border-box", 100D, 70D)]
    [InlineData("table", "", 100D, 70D)]
    [InlineData("td", "content-box", 130D, 130D)]
    [InlineData("td", "border-box", 100D, 100D)]
    [InlineData("td", "", 130D, 130D)]
    [InlineData("tr", "content-box", 100D, 100D)]
    [InlineData("tr", "border-box", 100D, 100D)]
    [InlineData("tr", "", 100D, 100D)]
    public void HtmlTableGeometry_HeightMinimaRespectBoxSizingAndInsets(string heightOwner, string sizing, double tableHeight, double cellHeight) {
        string Height(string tag) => tag == heightOwner ? "height:100px;padding:10px;border:5px solid black;"
            + (sizing.Length == 0 ? string.Empty : "box-sizing:" + sizing + ";") : "padding:0;";
        string html = TableGeometrySource("<table id='sized' style='" + Height("table") + "width:200px;margin:0;border-spacing:0;background:white'>"
            + "<tr style='" + Height("tr") + "'><td id='cell' style='" + Height("td") + "background:lime'>Sized</td></tr></table>"
            + "<div id='next' style='background:red;height:10px'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(tableHeight, TableGeometryShape(rendered, "table#sized").Height, 3);
        Assert.Equal(cellHeight, TableGeometryShape(rendered, "td#cell").Height, 3);
        Assert.Equal(tableHeight, TableGeometryShape(rendered, "div#next").Y, 3);
    }

    [Theory]
    [InlineData("top", 0D)]
    [InlineData("middle", 40D)]
    [InlineData("bottom", 80D)]
    public void HtmlTableGeometry_VerticalAlignPlacesContentInsideTheFinalRow(string alignment, double offset) {
        string html = TableGeometrySource("<table style='width:300px;margin:0;border-spacing:0'><tr>"
            + "<td style='padding:0'><div style='height:100px'>Tall</div></td>"
            + "<td id='aligned' style='padding:0;background:lime;vertical-align:" + alignment + "'>Aligned</td></tr></table>");

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderShape cell = TableGeometryShape(rendered, "td#aligned");
        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Aligned");

        Assert.Equal(100D, cell.Height, 3);
        Assert.Equal(cell.Y + offset, text.Y, 3);
    }

    [Theory]
    [InlineData("", "", "", 40D)]
    [InlineData("", "vertical-align:top", "", 0D)]
    [InlineData("", "vertical-align:bottom", "", 80D)]
    [InlineData("vertical-align:top", "", "", 0D)]
    [InlineData("vertical-align:bottom", "", "vertical-align:top", 0D)]
    [InlineData("", "vertical-align:bottom", "vertical-align:baseline", 0D)]
    [InlineData("", "", "vertical-align:initial", 0D)]
    [InlineData("", "", "vertical-align:unset", 0D)]
    public void HtmlTableGeometry_BrowserAlignmentDefaultsAndInheritanceRespectAuthoredOverrides(string groupStyle, string rowStyle, string cellStyle, double offset) {
        string html = TableGeometrySource("<table style='width:300px;margin:0;border-spacing:0'><tbody style='" + groupStyle
            + "'><tr style='" + rowStyle + "'><td style='padding:0'><div style='height:100px'>Tall</div></td>"
            + "<td id='aligned' style='padding:0;background:lime;" + cellStyle + "'>Aligned</td></tr></tbody></table>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderShape cell = TableGeometryShape(rendered, "td#aligned");
        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Aligned");

        Assert.Equal(cell.Y + offset, text.Y, 3);
    }

    [Theory]
    [InlineData("padding:0")]
    [InlineData("padding:initial")]
    [InlineData("padding-inline:0;padding-block:0")]
    public void HtmlTableGeometry_AuthoredZeroPaddingOverridesCellDefaults(string declaration) {
        string html = TableGeometrySource("<table style='width:200px;margin:0;border-spacing:0'><tr>"
            + "<td id='zero' style='" + declaration + ";background:lime'>Zero</td></tr></table>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderShape cell = TableGeometryShape(rendered, "td#zero");
        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>());

        Assert.Equal(cell.X, text.X, 3);
        Assert.Equal(cell.Y, text.Y, 3);
        Assert.Equal(20D, cell.Height, 3);
    }

    [Fact]
    public void HtmlTableGeometry_RowspanHeightAndBottomAlignmentKeepBothRowsAndFollowingContent() {
        const string body = "<table style='width:200px;margin:0;border-spacing:0'><tr>"
            + "<td id='spanned' rowspan='2' style='height:100px;padding:0;background:lime;vertical-align:bottom'>Spanned</td>"
            + "<td style='padding:0'>First</td></tr><tr><td style='padding:0'>Second</td></tr></table>"
            + "<div id='next' style='height:10px;background:red'>Next</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource(body), TableGeometryOptions());
        HtmlRenderShape cell = TableGeometryShape(rendered, "td#spanned");
        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Spanned");

        Assert.Equal(100D, cell.Height, 3);
        Assert.Equal(cell.Y + 80D, text.Y, 3);
        Assert.Equal(100D, TableGeometryShape(rendered, "div#next").Y, 3);
        Assert.Equal(new[] { "Spanned", "First", "Second", "Next" }, rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Select(item => item.Text));
    }

    [Fact]
    public void HtmlTableGeometry_RtlMirrorsSpannedColumnsWithoutReorderingLogicalContent() {
        const string body = "<table dir='rtl' style='width:240px;table-layout:fixed;margin:0;border-spacing:0'><tr>"
            + "<td id='first' rowspan='2' style='padding:0;background:lime'>First</td>"
            + "<td id='wide' colspan='2' style='padding:0;background:yellow'>Wide</td>"
            + "<td id='last' style='padding:0;background:orange'>Last</td></tr>"
            + "<tr><td colspan='3' style='padding:0'>SecondRow</td></tr></table>";
        string html = TableGeometrySource(body);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(180D, TableGeometryShape(rendered, "td#first").X, 3);
        Assert.Equal((60D, 120D), (TableGeometryShape(rendered, "td#wide").X, TableGeometryShape(rendered, "td#wide").Width));
        Assert.Equal(0D, TableGeometryShape(rendered, "td#last").X, 3);
        Assert.Equal(new[] { "First", "Wide", "Last", "SecondRow" }, rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Select(item => item.Text));
        PdfCore.PdfReadDocument pdf = PdfCore.PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions()));
        // Content spans follow the authored stream; plain-text extraction may
        // choose geometric left-to-right order for these English labels.
        string pdfText = string.Concat(pdf.Pages.SelectMany(page => page.GetTextSpans()).Select(span => span.Text));
        Assert.True(pdfText.IndexOf("First", StringComparison.Ordinal) < pdfText.IndexOf("Wide", StringComparison.Ordinal));
        Assert.True(pdfText.IndexOf("Wide", StringComparison.Ordinal) < pdfText.IndexOf("Last", StringComparison.Ordinal));
        Assert.Equal(4, pdf.TaggedContent!.StructureElements.Count(element => element.StructureType == "TD"));
    }

    [Theory]
    [InlineData("ltr")]
    [InlineData("rtl")]
    public void HtmlTableGeometry_CollapsedBorderColorTiesPreferTheLeadingSourceCell(string direction) {
        string html = TableGeometrySource("<table id='edges' dir='" + direction
            + "' style='width:200px;table-layout:fixed;margin:0;border-collapse:collapse'><tr>"
            + "<td style='padding:0;border:2px solid red'>First</td><td style='padding:0;border:2px solid blue'>Second</td></tr></table>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderShape shared = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "table#edges:collapsed-border-v-1-0");

        Assert.Equal(OfficeColor.Red, shared.Shape.StrokeColor);
        Assert.Equal(100D, shared.X, 3);
        HtmlRenderShape firstOuter = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "table#edges:collapsed-border-v-0-0");
        Assert.Equal(direction == "rtl" ? 200D : 0D, firstOuter.X, 3);
        Assert.Equal(OfficeColor.Red, firstOuter.Shape.StrokeColor);
    }

    [Theory]
    [InlineData("ltr", "right", "left")]
    [InlineData("rtl", "left", "right")]
    public void HtmlTableGeometry_CollapsedBorderTiesPreferLeadingCellBeforeTopmostSpanningCell(string direction, string leadingEdge, string trailingEdge) {
        string html = TableGeometrySource("<table id='edges' dir='" + direction
            + "' style='width:200px;table-layout:fixed;margin:0;border-collapse:collapse'><tr>"
            + "<td style='height:30px;padding:0'>First</td><td rowspan='2' style='padding:0;border-" + trailingEdge
            + ":2px solid blue'>Spanned</td></tr><tr><td style='height:30px;padding:0;border-" + leadingEdge
            + ":2px solid red'>Leading</td></tr></table>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderShape shared = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "table#edges:collapsed-border-v-1-1");

        Assert.Equal(OfficeColor.Red, shared.Shape.StrokeColor);
    }

    [Fact]
    public void HtmlTableGeometry_PagedAuthoredRowHeightsRepeatHeadersAndKeepBodyTextOnce() {
        string rows = string.Concat(Enumerable.Range(1, 5).Select(index =>
            "<tr style='height:40px'><td style='padding:0'>Body" + index + "</td></tr>"));
        string html = TableGeometrySource("<table style='width:200px;margin:0;border-spacing:0'>"
            + "<thead><tr style='height:20px'><th style='padding:0'>Header</th></tr></thead><tbody>" + rows + "</tbody></table>");
        var options = TableGeometryOptions();
        options.Mode = HtmlRenderMode.Paged;
        options.PageSize = new OfficePageSize(200D / 96D, 100D / 96D);
        options.HonorCssPageRules = false;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Header"));
        foreach (int index in Enumerable.Range(1, 5)) {
            Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text == "Body" + index);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
    }

    [Fact]
    public void HtmlTableGeometry_TableMinimumSurplusSurvivesPageContinuationReflow() {
        string rows = string.Concat(Enumerable.Range(1, 3).Select(index =>
            "<tr><td id='row" + index + "' style='padding:0;background:lime'>Body" + index + "</td></tr>"));
        string html = TableGeometrySource("<style>@page{size:200px 120px;margin:0}@page:left{size:220px 120px}</style>"
            + "<table style='width:180px;height:300px;margin:0;border-spacing:0'>" + rows + "</table>");
        var options = TableGeometryOptions();
        options.Mode = HtmlRenderMode.Paged;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(3, rendered.Pages.Count);
        foreach (int index in Enumerable.Range(1, 3)) {
            HtmlRenderShape cell = TableGeometryShape(rendered, "td#row" + index);
            Assert.Equal(100D, cell.Height, 3);
            Assert.Single(rendered.Pages[index - 1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Body" + index);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
    }

    [Theory]
    [InlineData("max-height", "min-content")]
    public void HtmlTableGeometry_IntrinsicBoxSizingFallbackCannotSatisfyNoLoss(string property, string value) {
        string html = TableGeometrySource("<style>.sized{--sizing:" + value + ";" + property + ":var(--sizing)}</style>"
            + "<div class='sized'>Intrinsic text</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);

        Assert.Contains(property + "=" + value, diagnostic.Detail);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
        var strict = TableGeometryOptions();
        strict.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
        Assert.Throws<HtmlConversionException>(() => HtmlRenderTestDriver.Render(html, strict));
    }

    [Fact]
    public void HtmlTableGeometry_DirectIntrinsicDeclarationsReportOneLossPerElement() {
        string html = TableGeometrySource("<div id='intrinsic' style='width:MAX-content;min-height:min-content;max-width:fit-content(100px)'>Text</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);

        Assert.Equal("div#intrinsic", diagnostic.Source);
        Assert.DoesNotContain("width=max-content", diagnostic.Detail);
        Assert.Contains("min-height=min-content", diagnostic.Detail);
        Assert.DoesNotContain("max-width=fit-content(100px)", diagnostic.Detail);
        Assert.True(rendered.HasLoss);
    }

    [Fact]
    public void HtmlTableGeometry_IntrinsicTrackSizingAndDiscardedDeclarationsDoNotReportBoxSizingLoss() {
        string html = TableGeometrySource("<div style='width:max-content;width:100px'>Authored override</div>"
            + "<div style='display:none;width:min-content'>Hidden</div>"
            + "<div style='display:grid;width:200px;grid-template-columns:max-content 1fr'><div>Track</div><div>Remaining</div></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
    }

    private static string TableGeometrySource(string body) =>
        "<!DOCTYPE html><style>html,body{margin:0;padding:0;font-size:16px;line-height:20px;font-family:Arial}</style>" + body;

    private static HtmlRenderOptions TableGeometryOptions() => new() {
        ViewportWidth = 600D,
        ViewportHeight = 300D,
        Margins = HtmlRenderMargins.All(0D),
        UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser
    };

    private static HtmlRenderShape TableGeometryShape(HtmlRenderDocument rendered, string source) =>
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(), shape => shape.Source == source && shape.Shape.FillColor.HasValue);
}
