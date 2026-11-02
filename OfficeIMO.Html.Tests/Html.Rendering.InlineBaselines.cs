using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlInlineText_SmallerDescendantsStayInsideTheirContainingLine() {
        const string html = "<div style='font-size:40px;line-height:48px;margin:0'>"
            + "<span id='small' style='font-size:10px;background:red'>Small</span></div>"
            + "<div style='margin:0'>Next block</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderText[] text = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText small = Assert.Single(text, item => item.Text == "Small");
        HtmlRenderText next = Assert.Single(text, item => item.Text.Contains("Next", StringComparison.Ordinal));
        var background = FindGridShape(rendered, "span#small");
        Assert.True(small.Y + small.Height <= next.Y + 0.001D);
        Assert.True(background.Y + background.Height <= next.Y + 0.001D);
    }

    [Theory]
    [InlineData("display:block", false)]
    [InlineData("display:flex;flex-direction:row;align-items:flex-start", false)]
    [InlineData("display:flex;flex-direction:column;align-items:flex-start", false)]
    [InlineData("display:block", true)]
    [InlineData("display:grid;grid-template-columns:150px 150px;grid-template-rows:60px;align-items:baseline", false)]
    public void HtmlInlineText_MixedSizesShareTheSavedPdfAndSvgBaseline(string containerStyle, bool inlineImage) {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            Margins = HtmlRenderMargins.All(0D),
            DefaultFontFamily = "Pinned"
        };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf")));
        string image = inlineImage ? "<img style='width:40px;height:40px' src='data:image/png;base64,"
            + Convert.ToBase64String(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(255, 0, 0)) + "'>" : "";
        string content = containerStyle.StartsWith("display:grid", StringComparison.Ordinal)
            ? "<span>AAAA</span><span style='font-size:10px;line-height:12px'>BBBB</span>"
            : "<span>AAAA <span style='font-size:10px;line-height:12px'>BBBB</span>" + image + "</span>";
        // Keep this alignment witness inside the page. Clipped text can have a
        // local SVG y plus a group transform, which is not its global baseline.
        var document = HtmlConversionDocument.Parse("<body style='margin:0;padding:24px'><div style='" + containerStyle
            + ";width:600px;font:40px/48px Pinned'>" + content + "</div></body>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        HtmlRenderText[] text = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText large = Assert.Single(text, node => node.Text.Trim() == "AAAA");
        HtmlRenderText small = Assert.Single(text, node => node.Text.Trim() == "BBBB");
        Assert.Equal(PaintBaseline(large), PaintBaseline(small), 3);
        AssertPdfBackedSvgRetainsBaselineInk(rendered.Pages[0], document.ToSvg(options));

        using var pdf = UglyToad.PdfPig.PdfDocument.Open((document.ToPdf(new HtmlToPdfOptions(options))).ToBytes());
        var letters = pdf.GetPage(1).Letters;
        double largeBaseline = letters.First(letter => letter.Value == "A").StartBaseLine.Y;
        double smallBaseline = letters.First(letter => letter.Value == "B").StartBaseLine.Y;
        Assert.Equal(largeBaseline, smallBaseline, 3);

    }
}
