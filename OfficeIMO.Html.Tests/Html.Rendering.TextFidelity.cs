using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlImageExport_UsesPixelSvgAndExactSharedTextAdvances() {
        const string html = "<html style='margin:0'><body style='margin:0'>"
            + "<p style='margin:0;font:14px Arial'><strong>One model.</strong> Shared output.</p>"
            + "<p dir='rtl' style='margin:0;font:14px Arial'>שלום 123</p>"
            + "</body></html>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 320D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D),
            Scale = 1.5D
        };

        HtmlRenderDocument rendered = html.RenderHtml(options);
        OfficeDrawing drawing = rendered.Pages[0].CreateDrawing();
        IReadOnlyList<HtmlRenderText> renderedText = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToList();
        IReadOnlyList<OfficeDrawingText> textElements = drawing.Elements.OfType<OfficeDrawingText>().ToList();
        string svg = html.ToSvg(options);
        byte[] png = html.ToPng(options);
        XNamespace ns = "http://www.w3.org/2000/svg";
        XDocument svgDocument = XDocument.Parse(svg);
        IReadOnlyList<XElement> svgText = svgDocument.Descendants(ns + "text").ToList();

        Assert.NotEmpty(textElements);
        Assert.Equal(renderedText.Count, textElements.Count);
        for (int index = 0; index < textElements.Count; index++) {
            OfficeDrawingText drawingText = textElements[index];
            HtmlRenderText sourceText = renderedText[index];
            Assert.Equal(OfficeTextOverflowBehavior.Clip, drawingText.OverflowBehavior);
            Assert.Equal(sourceText.Width, drawingText.Width, 6);
            Assert.Equal(
                Assert.IsType<double>(sourceText.TextAdvanceWidth),
                Assert.IsType<double>(drawingText.TextAdvanceWidth),
                6);
        }
        Assert.StartsWith("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"480px\" height=\"150px\" viewBox=\"0 0 320 100\"", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("pt\"", svg, StringComparison.Ordinal);
        Assert.All(svgText, text => {
            Assert.NotNull(text.Attribute("textLength"));
            Assert.Equal("spacingAndGlyphs", text.Attribute("lengthAdjust")?.Value);
        });
        Assert.Contains(svgText, text => text.Value == "One model.");
        Assert.Contains(svgText, text => text.Value.Contains("Shared output.", StringComparison.Ordinal));
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Take(8).ToArray());
    }

    [Fact]
    public void HtmlImageExport_ClipsTheTextFrameWithoutLosingItsResolvedAdvance() {
        const string html = "<html style='margin:0'><body style='margin:0'>"
            + "<div style='width:20px;white-space:nowrap;font-family:Arial;font-size:100px;line-height:100px'>WW</div>"
            + "</body></html>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 20D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = html.RenderHtml(options);
        HtmlRenderText text = rendered.Pages[0].Visuals
            .OfType<HtmlRenderText>()
            .First(item => item.TextAdvanceWidth.HasValue && item.TextAdvanceWidth.Value > item.Width);
        double advance = Assert.IsType<double>(text.TextAdvanceWidth);

        Assert.True(advance > text.Width);
        OfficeDrawing drawing = rendered.Pages[0].CreateDrawing();
        OfficeDrawingText drawingText = drawing.Elements
            .OfType<OfficeDrawingText>()
            .First(item => item.Text == text.Text && Math.Abs(item.X - text.X) < 0.000001D && Math.Abs(item.Y - text.Y) < 0.000001D);
        Assert.Equal(text.Width, drawingText.Width, 6);
        Assert.Equal(advance, Assert.IsType<double>(drawingText.TextAdvanceWidth), 6);
        Assert.NotEmpty(html.ToPng(options));
        Assert.StartsWith("<svg", html.ToSvg(options), StringComparison.Ordinal);
    }
}
