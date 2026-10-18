using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(286)]
    [InlineData(288)]
    [InlineData(300)]
    public void HtmlTables_ContentBreakRangesUseTheAvailableRowTailAfterAKeptParagraph(int pageHeight) {
        HtmlRenderDocument rendered = RenderContentRangeTable(ContentRangeTable(), pageHeight);

        Assert.Equal(2, rendered.Pages.Count);
        AssertContentRangeSource(rendered);
        AssertContentRangeLinesOnPage(rendered.Pages[0]);
        AssertContentRangeCellIdentity(rendered, 2);
        Assert.True(ContentRangeCells(rendered, "td#span-2").Count() > 1);
        HtmlRenderSemanticGroup footer = ContentRangeFooter(rendered.Pages[0]);
        Assert.Equal(rendered.Pages[0].Height - 20D, footer.Y + footer.Height, 3);

        if (pageHeight != 286) {
            HtmlRenderDocument baseline = RenderContentRangeTable(ContentRangeTable(), 286);
            Assert.Equal(pageHeight - 286D, footer.Y - ContentRangeFooter(baseline.Pages[0]).Y, 3);
        }
        HtmlRenderPage lastPage = rendered.Pages.Last();
        HtmlRenderText finalFooter = Assert.Single(lastPage.Visuals.OfType<HtmlRenderText>(), item => item.Text == "FOOTER");
        HtmlRenderText after = Assert.Single(lastPage.Visuals.OfType<HtmlRenderText>(), item => item.Text == "AFTER");
        Assert.True(after.Y >= finalFooter.Y + finalFooter.Height - 0.001D,
            $"AFTER starts at {after.Y}; the painted footer text occupies {finalFooter.Y}..{finalFooter.Y + finalFooter.Height}.");
    }

    [Fact]
    public void HtmlTables_ContentBreakRangesRetainTheFinalLineOfAnIndependentRow() {
        string table = "<table><thead><tr><th colspan='2'>HEADER</th></tr></thead>"
            + "<tbody><tr style='height:260px'><td id='span-2' class='span'>" + ContentRangeParagraph()
            + "</td><td>ROW-SIBLING</td></tr></tbody>" + ContentRangeFooterMarkup + "</table>";
        HtmlRenderDocument rendered = RenderContentRangeTable(table, 286);

        Assert.Equal(2, rendered.Pages.Count);
        AssertContentRangeMarkers(rendered, ContentRangeLines.Concat(new[] { "ROW-SIBLING", "AFTER" }));
        AssertContentRangeLinesOnPage(rendered.Pages[0]);
        AssertContentRangeCellIdentity(rendered, 1);
        HtmlRenderSemanticGroup footer = ContentRangeFooter(rendered.Pages[0]);
        Assert.Equal(rendered.Pages[0].Height - 20D, footer.Y + footer.Height, 3);
    }

    [Theory]
    [InlineData("block")]
    [InlineData("inline-link")]
    [InlineData("flex-column")]
    [InlineData("flex-row")]
    [InlineData("grid")]
    [InlineData("float")]
    [InlineData("multicolumn")]
    public void HtmlTables_ContentBreakRangesSurviveTheirFormattingOwner(string wrapper) {
        string table = ContentRangeTable(keptLink: wrapper != "inline-link");
        string body = wrapper switch {
            "inline-link" => "<a href='#after'>" + table + "</a>",
            "flex-column" => "<div style='display:flex;flex-direction:column;align-items:flex-start'>" + table + "</div>",
            "flex-row" => "<div style='display:flex;align-items:flex-start'><section style='width:300px'>"
                + table + "</section><aside style='width:40px'>SIDE</aside></div>",
            "grid" => "<div style='display:grid;grid-template-columns:300px 40px;align-items:start'><section>"
                + table + "</section><aside>SIDE</aside></div>",
            "float" => "<div style='display:flow-root'>" + table.Replace("<table>", "<table style='float:left'>")
                + "<p style='margin-left:310px'>SIDE</p></div>",
            "multicolumn" => "<div style='column-count:2;column-gap:60px;column-fill:balance'>" + table + "</div>",
            _ => "<div><section>" + table + "</section></div>"
        };
        HtmlRenderDocument rendered = RenderContentRangeTable(body, 286, pageWidth: wrapper == "multicolumn" ? 700 : 400);

        AssertContentRangeSource(rendered);
        AssertContentRangeCellIdentity(rendered, 2);
        HtmlRenderPage keptPage = Assert.Single(rendered.Pages, page => page.Visuals.OfType<HtmlRenderText>()
            .Any(item => item.Text == "LINE-A"));
        AssertContentRangeLinesOnPage(keptPage);
        if (wrapper != "multicolumn") Assert.Same(rendered.Pages[0], keptPage);
        if (wrapper is "flex-row" or "grid" or "float") AssertContentRangeMarkers(rendered, new[] { "SIDE" });
    }

    [Theory]
    [InlineData("row")]
    [InlineData("cell")]
    [InlineData("widows")]
    [InlineData("atomic-sibling")]
    public void HtmlTables_ContentBreakRangesRespectOverlappingCellConstraints(string constraint) {
        string css = constraint switch {
            "row" => ".pair-2>tr{break-inside:avoid}",
            "cell" => "#span-2{break-inside:avoid}",
            "widows" => ".kept{break-inside:auto;orphans:2;widows:3}",
            _ => ".kept{break-inside:auto}"
        };
        string sibling = constraint == "atomic-sibling"
            ? "FIRST-2<svg style='display:block' width='30' height='80'><rect width='30' height='80' fill='orange'/></svg>"
            : "FIRST-2";
        HtmlRenderDocument rendered = RenderContentRangeTable(ContentRangeTable(sibling: sibling), 246, css);

        AssertContentRangeSource(rendered);
        HtmlRenderPage keptPage = Assert.Single(rendered.Pages, page => page.Visuals.OfType<HtmlRenderText>()
            .Any(item => item.Text == "LINE-A"));
        Assert.NotSame(rendered.Pages[0], keptPage);
        AssertContentRangeLinesOnPage(keptPage);
        if (constraint == "atomic-sibling") {
            HtmlRenderDrawing drawing = Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
                .OfType<HtmlRenderDrawing>());
            Assert.Equal(80D, drawing.Height, 3);
            Assert.Contains(EnumerateRenderVisuals(keptPage.Scene), item => ReferenceEquals(item, drawing));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlTables_ContentBreakRangesKeepParagraphWidowsThroughColumns(bool columns) {
        string lines = string.Join("<br>", Enumerable.Range(1, 8).Select(i => "LONG-" + i));
        string table = "<table><tbody><tr style='height:150px'><td rowspan='2'><div style='line-height:20px;orphans:4;widows:4'>"
            + lines + "</div></td></tr><tr style='height:150px'></tr></tbody></table>";
        string body = columns ? "<div style='column-count:2;column-fill:auto;column-gap:60px;height:400px'>" + table + "</div>" : table;
        HtmlRenderDocument rendered = RenderContentRangeTable(body, 140, pageWidth: 700);

        AssertContentRangeMarkers(rendered, Enumerable.Range(1, 8).Select(i => "LONG-" + i));
        Assert.Equal(new[] { "LONG-1", "LONG-2", "LONG-3", "LONG-4" },
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Where(t => t.Text.StartsWith("LONG-", StringComparison.Ordinal)).Select(t => t.Text));
        Assert.Equal(new[] { "LONG-5", "LONG-6", "LONG-7", "LONG-8" },
            rendered.Pages[1].Visuals.OfType<HtmlRenderText>().Where(t => t.Text.StartsWith("LONG-", StringComparison.Ordinal)).Select(t => t.Text));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlTables_ContentBreakRangesCannotCutFollowingTextOverVisibleOverflow(bool inlineLink) {
        string labels = string.Join("<br>", Enumerable.Range(1, 20).Select(i => "OVER-" + i.ToString("D2")));
        string body = "<div style='display:flex;height:0;align-items:flex-start'><table><tr style='height:300px'><td>TABLE-TOP</td></tr></table></div>"
            + "<p style='padding-top:9px;line-height:20px;orphans:1;widows:1'>" + labels + "</p>";
        body = inlineLink ? "<a href='#after'>" + body + "</a>" : "<div>" + body + "</div>";
        HtmlRenderDocument rendered = RenderContentRangeTable(body, 140);

        AssertContentRangeMarkers(rendered, Enumerable.Range(1, 20).Select(i => "OVER-" + i.ToString("D2")).Append("TABLE-TOP"));
    }

    private static readonly string[] ContentRangeLines = { "LINE-A", "LINE-B", "LINE-C", "LINE-D" };
    private const string ContentRangeFooterMarkup = "<tfoot><tr><td id='footer' colspan='2'>FOOTER</td></tr></tfoot>";

    private static string ContentRangeParagraph(bool keptLink = true) {
        string paragraph = "<div class='kept'>" + string.Join("<br>", ContentRangeLines) + "</div>";
        return keptLink ? "<a href='#after'>" + paragraph + "</a>" : paragraph;
    }

    private static string ContentRangeTable(bool keptLink = true, string sibling = "FIRST-2") =>
        "<table><thead><tr><th colspan='2'>HEADER</th></tr></thead>"
        + "<tbody><tr><td id='span-1' class='span' rowspan='2'>SPAN-1</td><td>FIRST-1</td></tr>"
        + "<tr><td>SECOND-1</td></tr></tbody><tbody class='pair-2'><tr><td id='span-2' class='span' rowspan='2'>"
        + ContentRangeParagraph(keptLink) + "</td><td>" + sibling + "</td></tr><tr><td>SECOND-2</td></tr></tbody>"
        + ContentRangeFooterMarkup + "</table>";

    private static HtmlRenderDocument RenderContentRangeTable(string body, int pageHeight, string additionalCss = "", int pageWidth = 400) {
        string html = "<!doctype html><style>@page{size:" + pageWidth + "px " + pageHeight + "px;margin:20px}"
            + "html,body{margin:0;background:white;font:14px/18px RangeWitness}p{margin:0}"
            + "table{width:300px;table-layout:fixed;border-collapse:separate;border-spacing:0;background:#6bb7e6}"
            + "tbody>tr{height:60px}thead tr{height:28px}td,th{border:0;padding:0;orphans:1;widows:1;vertical-align:top}"
            + ".span{background:red}thead{background:#23395d;color:white}tfoot{background:#ddd}thead,tfoot{break-inside:avoid}"
            + ".kept{line-height:20px;break-inside:avoid;orphans:1;widows:1}"
            + additionalCss + "</style>" + body + "<p id='after'>AFTER</p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = pageWidth,
            ViewportHeight = pageHeight,
            Margins = HtmlRenderMargins.All(0D),
            DefaultFontFamily = "RangeWitness",
            AllowSystemFontFallback = false
        };
        string fonts = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Fonts");
        options.Fonts.Add("RangeWitness", File.ReadAllBytes(Path.Combine(fonts, "OfficeIMOBaselineSans-Regular.ttf")));
        options.Fonts.Add("RangeWitness", File.ReadAllBytes(Path.Combine(fonts, "OfficeIMOBaselineSans-Bold.ttf")), OfficeFontStyle.Bold);
        return HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html), HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.DisplayList, options)).Document;
    }

    private static void AssertContentRangeSource(HtmlRenderDocument rendered) =>
        AssertContentRangeMarkers(rendered, ContentRangeLines.Concat(new[] { "SPAN-1", "FIRST-1", "SECOND-1", "FIRST-2", "SECOND-2", "AFTER" }));

    private static void AssertContentRangeMarkers(HtmlRenderDocument rendered, IEnumerable<string> markers) {
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        foreach (string marker in markers) Assert.Single(text, item => item.Text == marker);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || item.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    private static void AssertContentRangeLinesOnPage(HtmlRenderPage page) {
        foreach (string line in ContentRangeLines) Assert.Single(page.Visuals.OfType<HtmlRenderText>(), item => item.Text == line);
    }

    private static IEnumerable<HtmlRenderSemanticGroup> ContentRangeCells(HtmlRenderDocument rendered, string source) =>
        rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderSemanticGroup>()
            .Where(group => group.Role == HtmlRenderSemanticGroupRole.TableCell && group.Source == source);

    private static void AssertContentRangeCellIdentity(HtmlRenderDocument rendered, int rowSpan) {
        HtmlRenderSemanticGroup[] cells = ContentRangeCells(rendered, "td#span-2").ToArray();
        Assert.NotEmpty(cells);
        Assert.All(cells, cell => Assert.Equal(rowSpan, cell.RowSpan));
        string? key = Assert.Single(cells.Select(cell => cell.StructureElementKey).Distinct());
        Assert.False(string.IsNullOrEmpty(key));
    }

    private static HtmlRenderSemanticGroup ContentRangeFooter(HtmlRenderPage page) =>
        Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderSemanticGroup>(),
            group => group.Role == HtmlRenderSemanticGroupRole.TableCell && group.Source == "td#footer");
}
