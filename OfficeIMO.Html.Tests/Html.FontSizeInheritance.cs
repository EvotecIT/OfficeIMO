using AngleSharp.Dom;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void HtmlFontSizesInheritResolvedValuesAndAnchorAbsoluteKeywordsToMedium() {
        const string html = """
            <div id="parent" style="font-size:150%">
              <p id="inherited" style="margin:0">Inherited</p>
              <p style="margin:0"><span id="nested"><b id="deeper">Nested</b></span></p>
              <p id="small" style="margin:0;font-size:small">Small</p>
              <p id="larger" style="margin:0;font-size:larger">Larger</p>
            </div>
            """;
        var parsed = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(parsed);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));

        Assert.Equal(18D, styles[parsed.QuerySelector("#parent")!].ResolvedFontSizePoints!.Value, 3);
        Assert.Equal(18D, styles[parsed.QuerySelector("#inherited")!].ResolvedFontSizePoints!.Value, 3);
        Assert.Equal(18D, styles[parsed.QuerySelector("#nested")!].ResolvedFontSizePoints!.Value, 3);
        Assert.Equal(18D, styles[parsed.QuerySelector("#deeper")!].ResolvedFontSizePoints!.Value, 3);
        Assert.Equal(10.68D, styles[parsed.QuerySelector("#small")!].ResolvedFontSizePoints!.Value, 3);
        Assert.Equal(21.6D, styles[parsed.QuerySelector("#larger")!].ResolvedFontSizePoints!.Value, 3);

        IReadOnlyList<HtmlRenderText> text = rendered.Pages.SelectMany(page => page.Visuals)
            .OfType<HtmlRenderText>().ToList();
        Assert.Equal(24D, Assert.Single(text, item => item.Text == "Inherited").Font.Size, 3);
        Assert.Equal(24D, Assert.Single(text, item => item.Text == "Nested").Font.Size, 3);
        Assert.Equal(14.24D, Assert.Single(text, item => item.Text == "Small").Font.Size, 3);
        Assert.Equal(28.8D, Assert.Single(text, item => item.Text == "Larger").Font.Size, 3);
    }
}
