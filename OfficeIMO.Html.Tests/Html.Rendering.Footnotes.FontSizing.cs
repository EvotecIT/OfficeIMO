using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "vertical-align:.5em", 24D, -5.6D)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "vertical-align:.5em", 24D, -5.6D)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "line-height:2;vertical-align:50%", 22.4D, -11.2D)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "line-height:2;vertical-align:50%", 22.4D, -11.2D)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "line-height:10px;vertical-align:50%", 10D, -5D)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "line-height:10px;vertical-align:50%", 10D, -5D)]
    public void HtmlFootnoteCall_DefaultFontResolvesRelativeAlignmentLikeAnAuthoredFont(
        HtmlRenderIntentProfile profile, string callStyle, double expectedLineHeight, double expectedOffset) {
        HtmlRenderText call = RenderCall(callStyle);
        HtmlRenderText control = RenderCall("font-size:70%;" + callStyle);

        Assert.Equal(11.2D, control.Font.Size, 6);
        Assert.Equal(expectedLineHeight, control.LineHeight, 6);
        Assert.Equal(expectedOffset, control.BaselineOffset, 6);
        Assert.Equal(control.Font.Size, call.Font.Size, 6);
        Assert.Equal(control.LineHeight, call.LineHeight, 6);
        Assert.Equal(control.BaselineOffset, call.BaselineOffset, 6);
        Assert.Equal(control.TextAdvanceWidth!.Value, call.TextAdvanceWidth!.Value, 6);
        Assert.Equal(control.LayoutY, call.LayoutY, 6);

        HtmlRenderText RenderCall(string style) {
            string html = "<style>body{margin:0;font:16px/24px Helvetica}.note{float:footnote}"
                + ".note::footnote-marker{content:none}.note::footnote-call{" + style + "}</style>"
                + "<p style='margin:0'>BASE<span id='note' class='note'>Note text</span> AFTER</p>";
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, options: new HtmlRenderOptions {
                    Margins = HtmlRenderMargins.All(0D)
                })).Document;
            return Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
                .OfType<HtmlRenderText>(), item => item.Source == "span#note:footnote-call");
        }
    }
}
