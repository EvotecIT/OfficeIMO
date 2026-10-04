using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, 0, 0D)]
    [InlineData(true, 0, 0D)]
    [InlineData(true, 16, 0D)]
    [InlineData(false, 0, 32D)]
    public void PagedRendererTruncatesUnforcedHeadingMarginsAndPreservesForcedMargins(
        bool column, int leadMargin, double expectedTargetY) {
        bool forced = expectedTargetY > 0D;
        string leadingLines = string.Join("<br>", Enumerable.Range(1, 17).Select(index => "Leading line " + index));
        string html = "<style>@page{size:400px 400px;margin:0}html,body{margin:0;padding:0}"
            + "body{font:16px/20px Arial}p{margin:0}#lead{margin-bottom:" + leadMargin + "px}"
            + "h3{font:20px/24px Arial;margin:32px 0 16px;break-after:avoid}"
            + (column ? "main{display:flex;flex-direction:column}" : "")
            + (forced ? "#target{break-before:page}" : "")
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
}
