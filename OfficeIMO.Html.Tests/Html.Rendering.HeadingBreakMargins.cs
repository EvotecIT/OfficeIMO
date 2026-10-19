using OfficeIMO.Html;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, 0, 0D, false)]
    [InlineData(false, 16, 0D, false)]
    [InlineData(true, 0, 0D, false)]
    [InlineData(true, 16, 0D, false)]
    [InlineData(false, 0, 32D, false)]
    [InlineData(false, 16, 32D, false)]
    [InlineData(true, 16, 0D, true)]
    [InlineData(true, 0, 32D, true)]
    public void PagedRendererTruncatesUnforcedHeadingMarginsAndPreservesForcedMargins(
        bool column, int leadMargin, double expectedTargetY, bool layeredBackground) {
        bool forced = expectedTargetY > 0D;
        string leadingLines = string.Join("<br>", Enumerable.Range(1, 17).Select(index => "Leading line " + index));
        string html = "<style>@page{size:400px 400px;margin:0}html,body{margin:0;padding:0}"
            + "body{font:16px/20px Arial}p{margin:0}#lead{margin-bottom:" + leadMargin + "px}"
            + "h3{font:20px/24px Arial;margin:32px 0 16px;break-after:avoid}"
            + (column ? "main{display:flex;flex-direction:column}" : "")
            + (forced ? "#target{break-before:page}" : "")
            + (layeredBackground ? "html{background:#222}body{background:#fff}" : "")
            + "</style><main><p id='lead'>" + leadingLines + "</p>"
            + "<h3 id='target'>Heading target</h3><p>Following one<br>Following two</p></main>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.Equal(2, rendered.Pages.Count);
        HtmlRenderVisual[] all = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).ToArray();
        foreach (string label in Enumerable.Range(1, 17).Select(index => "Leading line " + index)
            .Concat(new[] { "Heading target", "Following one", "Following two" })) {
            Assert.Single(all.OfType<HtmlRenderText>(), text => text.Text == label);
        }
        HtmlRenderBookmarkAnchor target = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderBookmarkAnchor>(),
            anchor => anchor.Source == "h3#target");
        Assert.Equal(expectedTargetY, target.Y, 3);
        if (layeredBackground) {
            foreach (HtmlRenderPage page in rendered.Pages) {
                Assert.Contains(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderShape>(),
                    shape => shape.Source == "body" && shape.Shape.FillColor == OfficeColor.White);
            }
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void PagedColumnKeepsHeadingWithFirstContentOfNestedColumn(bool padding, bool image) {
        string html = "<style>@page{size:400px 400px;margin:0}html,body{margin:0;padding:0}"
            + "body{font:16px/20px Arial}main,section{display:flex;flex-direction:column}"
            + "#lead{height:340px}h3{font:20px/24px Arial;margin:0;break-after:avoid}"
            + "section{" + (padding ? "padding-top" : "margin-top") + ":32px}p{margin:0}</style>"
            + "<main><div id='lead'>Lead</div><h3>Heading target</h3>"
            + "<section>" + (image
                ? "<svg id='following' style='width:20px;height:20px' viewBox='0 0 20 20' xmlns='http://www.w3.org/2000/svg'><rect width='20' height='20' fill='red'/></svg>"
                : "<p>Following line</p>") + "</section></main>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.Equal(2, rendered.Pages.Count);
        Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Text == "Heading target");
        foreach (string label in image ? new[] { "Heading target" } : new[] { "Heading target", "Following line" }) {
            Assert.Single(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>(),
                text => text.Text == label);
        }
        if (image) {
            HtmlRenderDrawing drawing = Assert.Single(
                EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderDrawing>());
            Assert.Equal(20D, drawing.Height, 3);
            Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderDrawing);
        }
    }

    [Theory]
    [InlineData("before")]
    [InlineData("after")]
    [InlineData("nested")]
    [InlineData("named")]
    public void PagedColumnRetainsFollowingMarginAtForcedBoundaries(string boundary) {
        string css = boundary switch {
            "before" or "nested" => "#target{break-before:page}",
            "after" => "#lead{break-after:page}",
            _ => "@page chapter{size:400px 500px}#target,#tail{page:chapter}"
        };
        string content = boundary == "nested"
            ? "<div id='lead' style='height:300px'>Lead</div><section><div style='height:40px'>Inner lead</div>"
                + "<h3 id='target'>Heading target</h3><p id='tail'>Following line</p></section>"
            : "<div id='lead' style='height:340px'>Lead</div>"
                + "<h3 id='target'>Heading target</h3><p id='tail'>Following line</p>";
        string html = "<style>@page{size:400px 400px;margin:0}html,body{margin:0;padding:0}"
            + "body{font:16px/20px Arial}main{display:flex;flex-direction:column}"
            + "h3{font:20px/24px Arial;margin:32px 0 16px;break-after:avoid}p{margin:0}"
            + css + "</style><main>" + content + "</main>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.Equal(2, rendered.Pages.Count);
        HtmlRenderBookmarkAnchor target = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderBookmarkAnchor>(),
            anchor => anchor.Source == "h3#target");
        Assert.Equal(32D, target.Y, 3);
        Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>(),
            text => text.Text == "Heading target");
        if (boundary == "named") Assert.Equal(500D, rendered.Pages[1].Height, 3);
    }

    [Theory]
    [InlineData("break-before:page", "", HtmlRenderMode.Paged, 1)]
    [InlineData("break-before:left", "", HtmlRenderMode.Paged, 1)]
    [InlineData("break-before:right", "", HtmlRenderMode.Paged, 2)]
    [InlineData("page-break-before:always", "", HtmlRenderMode.Paged, 1)]
    [InlineData("", "break-after:page", HtmlRenderMode.Paged, 1)]
    [InlineData("break-before:auto", "", HtmlRenderMode.Paged, 0)]
    [InlineData("break-before:avoid-page", "", HtmlRenderMode.Paged, 0)]
    [InlineData("break-before:page", "", HtmlRenderMode.Continuous, 0)]
    public void RendererKeepsFollowingChildMarginSeparateFromTableAtForcedPageBoundary(
        string before, string after, HtmlRenderMode mode, int targetPage) {
        string html = "<style>@page{size:400px 200px;margin:0}html,body{margin:0;padding:0}"
            + "body{font:16px/20px Arial}table{border-collapse:collapse;margin:0 0 16px;" + after + "}"
            + "td{padding:0}section{" + before + "}h2{font:20px/24px Arial;margin:18px 0 10px}"
            + "p{margin:0}</style><table id='lead'><tr><td>Table lead</td></tr></table>"
            + "<section><h2 id='target'>Heading target</h2><p>Following line</p></section>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = mode, ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) });

        Assert.Equal(targetPage + 1, rendered.Pages.Count);
        HtmlRenderSemanticGroup table = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "table#lead" && group.Role == HtmlRenderSemanticGroupRole.Table);
        Assert.Equal(0D, table.Y, 3);
        HtmlRenderBookmarkAnchor target = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[targetPage].Scene).OfType<HtmlRenderBookmarkAnchor>(),
            anchor => anchor.Source == "h2#target");
        Assert.Equal(targetPage > 0 ? 18D : table.Y + table.Height + 18D, target.Y, 3);
        foreach (string text in new[] { "Heading target", "Following line" }) {
            Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>(),
                item => item.Text == text);
            Assert.Single(EnumerateRenderVisuals(rendered.Pages[targetPage].Scene).OfType<HtmlRenderText>(),
                item => item.Text == text);
        }
    }

    [Theory]
    [InlineData("heading", false)]
    [InlineData("wrapper", false)]
    [InlineData("named", false)]
    [InlineData("heading", true)]
    [InlineData("wrapper", true)]
    [InlineData("named", true)]
    [InlineData("wrapped-single-column", false)]
    [InlineData("interrupted", false)]
    public void RendererKeepsFollowingMarginAtResolvedPageBoundary(string boundary, bool layeredBackground) {
        string headingBreak = boundary == "named" ? "" : "break-before:page;";
        string sectionPage = boundary == "named" ? "page:chapter;" : "";
        string content = "<h2 id='target'>Heading target</h2><p>Following line</p>";
        if (boundary == "wrapper") content = "<div>" + content + "</div>";
        if (boundary == "wrapped-single-column") content = "<div style='display:flex;flex-flow:column wrap;height:150px'><section>" + content + "</section></div>";
        if (boundary == "interrupted") content = "<span><p style='margin-bottom:16px'>Inner lead</p>" + content + "</span>";
        string html = "<style>@page{size:400px 200px;margin:0}@page chapter{size:400px 200px;margin:0}"
            + "html,body{margin:0;padding:0}body{font:16px/20px Arial}"
            + "table{border-collapse:collapse;margin:0 0 16px}td{padding:0}"
            + "section{" + sectionPage + "}h2{font:20px/24px Arial;margin:18px 0 10px;" + headingBreak + "}"
            + "p{margin:0}" + (layeredBackground ? "html{background:#222}body{background:white}" : "")
            + "</style><table><tr><td>Table lead</td></tr></table><section>" + content + "</section>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) });

        Assert.Equal(2, rendered.Pages.Count);
        HtmlRenderBookmarkAnchor target = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderBookmarkAnchor>(),
            anchor => anchor.Source == "h2#target");
        Assert.Equal(18D, target.Y, 3);
        foreach (string text in new[] { "Heading target", "Following line" }) {
            Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>(),
                item => item.Text == text);
        }
    }

    [Theory]
    [InlineData("columns")]
    [InlineData("table")]
    [InlineData("nested-table")]
    [InlineData("row-flex")]
    [InlineData("grid")]
    [InlineData("reverse-column")]
    [InlineData("wrapped-columns")]
    [InlineData("nested-table-flex")]
    [InlineData("inline-block")]
    [InlineData("inline-flex")]
    [InlineData("inline-grid")]
    [InlineData("positioned")]
    [InlineData("page-float")]
    [InlineData("footnote")]
    public void RendererKeepsCollapsedMarginsWhenPageControlUsesSourceFlowFallback(string context) {
        HtmlRenderDocument RenderWithBreak(string before) {
            string content = "<p style='margin:0 0 16px'>Lead</p><p id='target' style='margin:18px 0 0;break-before:" + before + "'>Target</p>";
            if (context == "nested-table") content = "<div>" + content + "</div>";
            if (context == "nested-table-flex") content = "<div style='display:flex;flex-direction:column'><div>" + content + "</div></div>";
            string wrapper = context switch {
                "columns" => "<div style='columns:2;column-gap:20px;column-fill:auto;height:100px'>" + content + "</div>",
                "row-flex" => "<div style='display:flex'><div>" + content + "</div></div>",
                "grid" => "<div style='display:grid'><div>" + content + "</div></div>",
                "reverse-column" => "<div style='display:flex;flex-direction:column-reverse'><div>" + content + "</div></div>",
                "wrapped-columns" => "<div style='display:flex;flex-flow:column wrap;height:50px'><div>" + content + "</div><div style='height:10px'>Filler</div></div>",
                "inline-block" => "<span style='display:inline-block;width:200px'>" + content + "</span>",
                "inline-flex" => "<span style='display:inline-flex;flex-direction:column;width:200px'><div>" + content + "</div></span>",
                "inline-grid" => "<span style='display:inline-grid;width:200px'><div>" + content + "</div></span>",
                "positioned" => "<div style='position:absolute;left:0;top:0;width:200px'>" + content + "</div>",
                "page-float" => "<div style='float:top;float-reference:page;width:200px'>" + content + "</div>",
                "footnote" => "<div style='float:footnote'>" + content + "</div>",
                _ => "<table><tr><td>" + content + "</td></tr></table>"
            };
            string html = "<style>@page{size:400px 200px;margin:0}html,body{margin:0;padding:0}"
                + "body{font:16px/20px Arial}table{border-collapse:collapse;margin:0}td{padding:0}</style>" + wrapper;
            return HtmlRenderTestDriver.Render(html,
                new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) });
        }

        HtmlRenderDocument ordinary = RenderWithBreak("auto");
        HtmlRenderDocument hinted = RenderWithBreak("page");
        HtmlRenderText Text(HtmlRenderDocument document, string value) => Assert.Single(
            EnumerateRenderVisuals(Assert.Single(document.Pages).Scene).OfType<HtmlRenderText>(),
            text => text.Text == value);
        HtmlRenderText expected = Text(ordinary, "Target");
        HtmlRenderText actual = Text(hinted, "Target");
        Assert.Equal(38D, expected.Y - Text(ordinary, "Lead").Y, 3);
        Assert.Equal(expected.X, actual.X, 3);
        Assert.Equal(expected.Y, actual.Y, 3);
    }

    [Fact]
    public void PagedWrappedSingleColumnMeasuresFollowingItemAfterResolvedMarginBoundary() {
        string html = "<style>@page{size:400px 200px;margin:0}html,body{margin:0;padding:0}"
            + "body{font:16px/20px Arial}main{display:flex;flex-flow:column wrap;height:150px}p{margin:0}"
            + "</style><main><div><p style='margin-bottom:16px'>Lead</p>"
            + "<p style='margin-top:18px;break-before:page'>Target</p></div><p>Tail</p></main>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.Equal(2, rendered.Pages.Count);
        HtmlRenderText[] texts = EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText target = Assert.Single(texts, item => item.Text == "Target");
        HtmlRenderText tail = Assert.Single(texts, item => item.Text == "Tail");
        Assert.Equal(20D, tail.Y - target.Y, 3);
        Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(), item => item.Text == "Lead");
    }

    [Fact]
    public void PagedMarginBoundaryRelayoutDoesNotRetainAbandonedFloatPlacements() {
        HtmlRenderDocument Render(int leadMargin) => HtmlRenderTestDriver.Render(
            "<style>@page{size:400px 200px;margin:0}html,body{margin:0;padding:0}body{font:16px/20px Arial}"
            + "p{margin:0}#lead{margin-bottom:" + leadMargin + "px}section{margin-top:18px;break-before:page}"
            + "</style><p id='lead'>Lead</p><section><div style='float:left;width:100px;height:40px'>Float</div>"
            + "<p>Target</p></section>", new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderDocument control = Render(0);
        HtmlRenderDocument collapsed = Render(16);
        Assert.Equal(2, control.Pages.Count);
        Assert.Equal(2, collapsed.Pages.Count);
        foreach (string value in new[] { "Float", "Target" }) {
            HtmlRenderText expected = Assert.Single(EnumerateRenderVisuals(control.Pages[1].Scene).OfType<HtmlRenderText>(), item => item.Text == value);
            HtmlRenderText actual = Assert.Single(EnumerateRenderVisuals(collapsed.Pages[1].Scene).OfType<HtmlRenderText>(), item => item.Text == value);
            Assert.Equal(expected.X, actual.X, 3);
            Assert.Equal(expected.Y, actual.Y, 3);
        }
    }
}
