using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("collapse")]
    [InlineData("separate")]
    public void HtmlTables_PagedBackgroundLayersAndOwnSidesSurviveSafeRowCuts(string borderModel) {
        HtmlRenderDocument rendered = RenderTableFragmentBackground(borderModel);
        Assert.Equal(2, rendered.Pages.Count);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        Assert.Equal(Enumerable.Range(1, 7).Select(index => $"R{index:00}"),
            text.Where(item => item.Text.Length == 3 && item.Text.StartsWith("R", StringComparison.Ordinal))
                .Select(item => item.Text));
        Assert.Equal(2, text.Count(item => item.Text == "HEAD-A"));
        Assert.Single(text, item => item.Text == "AFTER-TABLE");

        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        OfficeRasterImage last = OfficeDrawingRasterRenderer.Render(rendered.Pages[1].CreateDrawing(), 1D, OfficeColor.White);
        OfficeColor table = OfficeColor.FromRgb(0x6b, 0xb7, 0xe6);
        OfficeColor canvas = OfficeColor.FromRgb(0xee, 0xf2, 0xf6);
        OfficeColor border = OfficeColor.FromRgb(0x17, 0x20, 0x33);

        // The native fixture retains table paint through slack, while its row/group/cell
        // layers and collapsed grid sides stop at the last complete row.
        Assert.Equal(table, first.GetPixel(25, 280));
        Assert.Equal(table, first.GetPixel(345, 280));
        foreach (int x in new[] { 20, 21, 354, 355 }) {
            Assert.Equal(borderModel == "separate" ? border : table, first.GetPixel(x, 280));
            Assert.Equal(canvas, last.GetPixel(x, 280));
        }
        Assert.Equal(table, first.GetPixel(180, 299));
        Assert.Equal(canvas, first.GetPixel(380, 280));
        Assert.Equal(canvas, last.GetPixel(25, 280));

        Assert.Equal(table, first.GetPixel(25, 70));
        Assert.Equal(OfficeColor.FromRgb(0xa9, 0xd7, 0xa0), first.GetPixel(25, 120));
        Assert.Equal(OfficeColor.FromRgb(0xf5, 0xd4, 0x7a), first.GetPixel(25, 175));
        foreach (int y in new[] { 70, 120, 175, 225 }) {
            Assert.Equal(OfficeColor.FromRgb(0xe7, 0xa7, 0xcf), first.GetPixel(345, y));
        }
    }

    [Fact]
    public void HtmlTables_ContinuedSolidOuterSidesKeepTheirOwnColors() {
        HtmlRenderDocument rendered = RenderTableFragmentBackground("separate",
            "border-left-color:red;border-right-color:lime;");
        Assert.Equal(2, rendered.Pages.Count);
        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.Red, first.GetPixel(20, 280));
        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), first.GetPixel(355, 280));
        Assert.Equal(OfficeColor.FromRgb(0x6b, 0xb7, 0xe6), first.GetPixel(180, 280));
    }

    [Fact]
    public void HtmlRender_BorderedOrdinaryBoxDoesNotAcquireTableSlackPaint() {
        const string html = """
            <style>
            @page { size:240px 240px; margin:0 }
            html,body { margin:0; font:14px/18px Arial; background:white }
            .lead { height:100px }
            .panel { background:blue; border:2px solid red }
            .start { height:40px }
            .kept { height:140px; break-inside:avoid }
            </style>
            <div class='lead'>LEADING</div>
            <div class='panel'><div class='start'>START</div><div class='kept'>KEPT</div></div>
            <div>AFTER</div>
            """;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true,
            ViewportWidth = 240D, ViewportHeight = 240D, Margins = HtmlRenderMargins.All(0D)
        });
        Assert.Equal(2, rendered.Pages.Count);
        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.Blue, first.GetPixel(100, 120));
        Assert.Equal(OfficeColor.White, first.GetPixel(100, 200));
        Assert.Equal(OfficeColor.White, first.GetPixel(0, 200));
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        foreach (string label in new[] { "LEADING", "START", "KEPT", "AFTER" })
            Assert.Single(text, item => item.Text == label);
    }

    private static HtmlRenderDocument RenderTableFragmentBackground(string borderModel, string tableStyle = "") {
        string Row(int index) => "<tr" + (index is 3 or 6 ? " class='row-color'" : string.Empty)
            + "><td>R" + index.ToString("00") + "</td><td>BASE-" + index.ToString("00")
            + "</td><td>ROW-" + index.ToString("00") + "</td><td class='cell-color'>CELL-"
            + index.ToString("00") + "</td></tr>";
        string html = "<style>@page{size:400px 320px;margin:20px}html{background:#eef2f6}"
            + "body{margin:0;color:#172033;font:14px/18px Arial}"
            + "table{width:336px;table-layout:fixed;border-collapse:" + borderModel
            + ";border-spacing:0;border:2px solid #172033;background:#6bb7e6;" + tableStyle + "}"
            + "thead{display:table-header-group;background:#23395d;color:white}"
            + "tbody.group-color{background:#a9d7a0}tr{break-inside:avoid;page-break-inside:avoid}"
            + "th,td{box-sizing:border-box;border:2px solid #172033;padding:0 6px;text-align:left;vertical-align:middle}"
            + "th{height:28px;font-weight:normal}td{height:52px}tr.row-color{background:#f5d47a}"
            + "td.cell-color{background:#e7a7cf}p{margin:8px 0 0}</style>"
            + "<table><thead><tr><th>HEAD-A</th><th>HEAD-B</th><th>HEAD-C</th><th>HEAD-D</th></tr></thead>"
            + "<tbody>" + Row(1) + "</tbody><tbody class='group-color'>"
            + string.Concat(Enumerable.Range(2, 6).Select(Row)) + "</tbody></table><p>AFTER-TABLE</p>";
        return HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true,
            ViewportWidth = 400D, ViewportHeight = 320D, Margins = HtmlRenderMargins.All(0D)
        });
    }
}
