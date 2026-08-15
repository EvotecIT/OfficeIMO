using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("block", false)]
    [InlineData("block", true)]
    [InlineData("flex", false)]
    [InlineData("flex", true)]
    [InlineData("flex;flex-direction:column", false)]
    [InlineData("flex;flex-wrap:wrap", true)]
    [InlineData("grid", false)]
    [InlineData("grid", true)]
    public void HtmlRender_Paged_MixedOrdinaryAndCapturedRunningStringsUseDomOrder(string layout, bool capturedFirst) {
        const string ordinary = "<span style=\"string-set:title 'Ordinary'\">Body</span>";
        const string captured = "<span class=\"running\">Header</span>";
        string children = capturedFirst ? captured + ordinary : ordinary + captured;
        string html = """
            <style>
              @page {
                size:320px 180px;
                margin:32px;
                @top-left { content:string(title, first); }
                @top-right { content:string(title, last); }
              }
              .container { display:LAYOUT; width:240px; }
              .running { position:running(header); string-set:title 'Captured'; }
            </style>
            <div class="container">
              CHILDREN
            </div>
            """.Replace("LAYOUT", layout).Replace("CHILDREN", children);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        });
        IReadOnlyList<HtmlRenderText> marginText = Assert.Single(rendered.Pages).Visuals
            .OfType<HtmlRenderText>()
            .Where(text => text.SemanticRole == "page-margin")
            .OrderBy(text => text.X)
            .ToList();

        Assert.Equal(2, marginText.Count);
        Assert.Equal(capturedFirst ? "Captured" : "Ordinary", marginText[0].Text);
        Assert.Equal(capturedFirst ? "Ordinary" : "Captured", marginText[1].Text);
        Assert.Empty(rendered.Diagnostics);
    }
}
