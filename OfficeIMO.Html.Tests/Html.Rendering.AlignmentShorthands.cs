using AngleSharp.Html.Parser;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlAlignmentShorthandTests {
    [Theory]
    [InlineData("align-items:center", "center", "")]
    [InlineData("justify-items:end", "", "end")]
    [InlineData("place-items:center end", "center", "end")]
    [InlineData("place-items:center;justify-items:start", "center", "start")]
    [InlineData("justify-items:start;place-items:center", "center", "center")]
    [InlineData("justify-items:start!important;place-items:center", "center", "start")]
    [InlineData("--place:center end;place-items:var(--place);justify-items:start", "center", "start")]
    [InlineData("place-items:center;place-items:initial", "", "")]
    public void AlignmentShorthandKeepsAxesAndCascadeIndependent(string declaration, string align, string justify) {
        foreach (bool inline in new[] { true, false }) {
            string html = (inline ? "" : "<style>#target{" + declaration + "}</style>")
                + "<div id='target'" + (inline ? " style='" + declaration + "'" : "") + ">Text</div>";
            var document = new HtmlParser().ParseDocument(html);
            HtmlComputedStyle computed = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];
            Assert.Equal(align, computed.GetValue("align-items"));
            Assert.Equal(justify, computed.GetValue("justify-items"));
        }
    }
}

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlGrid_VerticalCenteringKeepsShortAndLongRowsAtTheSameStart() {
        string html = "<style>.row{display:grid;width:300px;grid-template-columns:200px 100px;align-items:center}</style>"
            + "<div class='row'><div id='short' style='height:10px;background:red'>Short</div><div>Metric</div></div>"
            + "<div class='row'><div id='long' style='height:10px;background:blue'>A longer title</div><div>Metric</div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 340D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape shortRow = FindGridShape(rendered, "div#short");
        HtmlRenderShape longRow = FindGridShape(rendered, "div#long");
        Assert.Equal(0D, shortRow.X, 3);
        Assert.Equal(200D, shortRow.Width, 3);
        Assert.Equal(shortRow.X, longRow.X, 3);
    }
}
