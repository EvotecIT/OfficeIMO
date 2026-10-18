using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("collapse", "top", HtmlRenderIntentProfile.PrintPaged)]
    [InlineData("separate", "top", HtmlRenderIntentProfile.PrintPaged)]
    [InlineData("collapse", "bottom", HtmlRenderIntentProfile.PrintPaged)]
    [InlineData("collapse", "top", HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlTables_RowSpanAllowsSafeRowCutsWithoutAnAuthoredCellKeep(
        string borderModel, string alignment, HtmlRenderIntentProfile profile) {
        HtmlRenderDocument rendered = RenderRowSpanPagination(borderModel, alignment: alignment, profile: profile);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "FIRST-2");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "SECOND-2");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "SECOND-2");
        AssertRowSpanPaginationContent(rendered, "SPAN-1", "SPAN-2");
        foreach (HtmlRenderPage page in rendered.Pages) {
            Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "HEADER");
            Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "FOOTER");
        }
        HtmlRenderSemanticGroup[] spans = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderSemanticGroup>().Where(group => group.Source == "td#span-2").ToArray();
        Assert.All(rendered.Pages, page => Assert.Contains(EnumerateRenderVisuals(page.Scene),
            visual => visual is HtmlRenderSemanticGroup { Source: "td#span-2" }));
        Assert.All(spans, span => Assert.Equal(2, span.RowSpan));
        Assert.Single(spans.Select(span => span.StructureElementKey).Distinct());
    }

    [Theory]
    [InlineData("td.span{break-inside:avoid}", "top")]
    [InlineData("tbody{break-inside:avoid}", "top")]
    [InlineData("tbody>tr{break-inside:avoid}", "top")]
    [InlineData("", "middle")]
    public void HtmlTables_RowSpanProtectsAuthoredKeepsAndCutsThroughText(string additionalCss, string alignment) {
        HtmlRenderDocument rendered = RenderRowSpanPagination("collapse", additionalCss: additionalCss, alignment: alignment);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "FIRST-2");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "FIRST-2");
        AssertRowSpanPaginationContent(rendered, "SPAN-1", "SPAN-2");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlTables_RowSpanRetainsNestedBlockKeepsAtRowCuts(bool keep) {
        const string lines = "LINE-A<br>LINE-B<br>LINE-C<br>LINE-D";
        HtmlRenderDocument rendered = RenderRowSpanPagination("separate", secondSpan:
            "<div style='line-height:20px;orphans:1;widows:1;" + (keep ? "break-inside:avoid" : string.Empty) + "'>" + lines + "</div>",
            additionalCss: "td,th{border:0;padding:0}");

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(!keep, rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Any(text => text.Text == "FIRST-2"));
        AssertRowSpanPaginationContent(rendered, "SPAN-1", "LINE-A", "LINE-B", "LINE-C", "LINE-D");
    }

    [Fact]
    public void HtmlTables_RowSpanHonorsParagraphWidowsAtAnAlignedRowCut() {
        HtmlRenderDocument rendered = RenderRowSpanPagination("separate", secondSpan:
            "<div style='line-height:20px;orphans:2;widows:2'>LINE-A<br>LINE-B<br>LINE-C<br>LINE-D</div>",
            additionalCss: "td,th{border:0;padding:0}");

        Assert.Equal(2, rendered.Pages.Count);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "FIRST-2");
        AssertRowSpanPaginationContent(rendered, "SPAN-1", "LINE-A", "LINE-B", "LINE-C", "LINE-D");
    }

    [Fact]
    public void HtmlTables_RowSpanKeepsItsOriginWhenFollowingPageGeometryChanges() {
        HtmlRenderDocument rendered = RenderRowSpanPagination("collapse", secondSpan:
            "<a href='#after'>SPAN-2</a>", additionalCss: "@page:left{size:340px 278px}");

        Assert.Equal(2, rendered.Pages.Count);
        AssertRowSpanPaginationContent(rendered, "SPAN-1", "SPAN-2");
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "FIRST-2");
    }

    private static void AssertRowSpanPaginationContent(HtmlRenderDocument rendered, params string[] spanLabels) {
        foreach (string label in spanLabels.Concat(new[] { "FIRST-1", "SECOND-1", "FIRST-2", "SECOND-2", "AFTER" })) {
            Assert.Single(rendered.Pages.SelectMany(page => page.Visuals.OfType<HtmlRenderText>()), text => text.Text == label);
        }
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || item.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    private static HtmlRenderDocument RenderRowSpanPagination(string borderModel, string alignment = "top",
        string additionalCss = "", string secondSpan = "SPAN-2", HtmlRenderIntentProfile profile = HtmlRenderIntentProfile.PrintPaged) {
        string html = "<style>@page{size:400px " + (borderModel == "separate" ? 286 : 278) + "px;margin:20px}html,body{margin:0;font:14px/18px Arial}"
            + "table{width:300px;table-layout:fixed;border-collapse:" + borderModel + ";border-spacing:0;background:#6bb7e6}"
            + "tbody>tr{height:60px}td,th{border:2px solid #172033;padding:4px;vertical-align:" + alignment + "}"
            + ".span{background:red}thead tr{height:28px}thead,tfoot{break-inside:avoid}thead{background:#23395d;color:white}tfoot{background:#ddd}p{margin:0}"
            + additionalCss + "</style><table><thead><tr><th colspan='2'>HEADER</th></tr></thead>"
            + "<tbody><tr><td class='span' rowspan='2'>SPAN-1</td><td>FIRST-1</td></tr><tr><td>SECOND-1</td></tr></tbody>"
            + "<tbody><tr><td id='span-2' class='span' rowspan='2'>" + secondSpan + "</td><td>FIRST-2</td></tr>"
            + "<tr><td>SECOND-2</td></tr></tbody><tfoot><tr><td colspan='2'>FOOTER</td></tr></tfoot></table><p id='after'>AFTER</p>";
        return HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html), HtmlRenderRequest.Create(profile,
            HtmlRenderEncoder.DisplayList, new HtmlRenderOptions { ViewportWidth = 400D, ViewportHeight = 320D,
                Margins = HtmlRenderMargins.All(0D) })).Document;
    }
}
