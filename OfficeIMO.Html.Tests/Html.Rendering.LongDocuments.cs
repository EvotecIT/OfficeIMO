using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_NestedBreakBeforeStartsTheDescendantOnANewPage() {
        const string html = "<main><p>First marker</p><section style='break-before:page'>Second marker</section></main>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, PageSize = new OfficePageSize(4D, 3D) };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains("First marker", GetPageText(rendered.Pages[0]), StringComparison.Ordinal);
        Assert.Contains("Second marker", GetPageText(rendered.Pages[1]), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_InternalForcedBreakUsesTheRemainingCurrentPageBeforeBreaking() {
        const string html = "<div style='height:120px'>Prefix marker</div><main><div style='height:80px'>Before marker</div><div style='height:80px;break-before:page'>After marker</div></main>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains("Prefix marker", GetPageText(rendered.Pages[0]), StringComparison.Ordinal);
        Assert.Contains("Before marker", GetPageText(rendered.Pages[0]), StringComparison.Ordinal);
        Assert.Contains("After marker", GetPageText(rendered.Pages[1]), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_NestedRightPageBreakInsertsTheRequiredBlankLeftPage() {
        const string html = "<main><p>First marker</p><section style='break-before:right'>Right-page marker</section></main>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, PageSize = new OfficePageSize(4D, 3D) };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.Contains("First marker", GetPageText(rendered.Pages[0]), StringComparison.Ordinal);
        Assert.DoesNotContain("Right-page marker", GetPageText(rendered.Pages[1]), StringComparison.Ordinal);
        Assert.Contains("Right-page marker", GetPageText(rendered.Pages[2]), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_FirstDescendantBreakBeforeWinsAtACollapsedParentBoundary() {
        const string html = "<p>First marker</p><main style='break-before:left'><section style='break-before:right'>Right-page marker</section></main>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, PageSize = new OfficePageSize(4D, 3D) };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.Contains("First marker", GetPageText(rendered.Pages[0]), StringComparison.Ordinal);
        Assert.DoesNotContain("Right-page marker", GetPageText(rendered.Pages[1]), StringComparison.Ordinal);
        Assert.Contains("Right-page marker", GetPageText(rendered.Pages[2]), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_NestedTrailingBreakAfterDoesNotCreateABlankFinalPage() {
        const string html = "<main><section style='break-after:page'>Only marker</section></main>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, PageSize = new OfficePageSize(4D, 3D) };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        HtmlRenderPage page = Assert.Single(rendered.Pages);
        Assert.Contains("Only marker", GetPageText(page), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_LongPagedDocumentPreservesEveryPageMarkerExactlyOnce() {
        const int pageCount = 100;
        var html = new StringBuilder(pageCount * 320 + 1024);
        html.Append("<style>@page{size:letter;margin:.65in;@bottom-right{content:'Page ' counter(page) ' of ' counter(pages)}}")
            .Append("body{margin:0;font:11px/1.4 Arial}section{break-after:page}section:last-child{break-after:auto}</style><main>");
        for (int index = 0; index < pageCount; index++) {
            html.Append("<section><h1>Article ")
                .Append(index + 1)
                .Append("</h1><p>Document marker PAGE-")
                .Append(index.ToString("D4"))
                .Append(".</p><p>Deterministic legal packet content for pagination verification.</p></section>");
        }
        html.Append("</main>");

        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = OfficePageSizes.Letter,
            Margins = HtmlRenderMargins.All(48D),
            MaxPageCount = pageCount
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html.ToString(), options);

        Assert.Equal(pageCount, rendered.Pages.Count);
        for (int index = 0; index < pageCount; index++) {
            string marker = "PAGE-" + index.ToString("D4");
            Assert.Equal(1, rendered.Text.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
            Assert.Contains(rendered.Pages[index].Visuals.OfType<HtmlRenderText>(), text =>
                text.Text == "Page " + (index + 1) + " of " + pageCount);
        }
    }

    private static string GetPageText(HtmlRenderPage page) => string.Concat(
        EnumerateRenderVisuals(page.Visuals)
            .OfType<HtmlRenderText>()
            .Select(text => text.Text));
}
