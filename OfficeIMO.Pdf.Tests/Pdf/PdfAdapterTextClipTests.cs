using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAdapterTextClipTests {
    [Theory]
    [InlineData("font:32px Arial;line-height:1")]
    [InlineData("font:bold 32px Arial;line-height:1.2")]
    [InlineData("font:italic 32px Arial;line-height:.8")]
    public void HtmlLineMetricsDoNotIntroduceAnUnauthoredGlyphClip(string style) {
        byte[] bytes = HtmlConversionDocument.Parse("<p style='margin:0;" + style + "'>gypqj</p>")
            .ToPdfBytes(new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0) });

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetTextSpans());
        Assert.Equal("gypqj", span.Text);
        Assert.Null(span.ClipPath);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlOverflowStillClipsTextAndLinkAnnotations(bool link) {
        string text = link ? "<a href='https://example.com'>gypqj</a>" : "gypqj";
        byte[] bytes = HtmlConversionDocument.Parse(
            "<div style='width:36px;height:12px;overflow:hidden;font:32px Arial'>" + text + "</div>")
            .ToPdfBytes(new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0) });

        var page = PdfReadDocument.Open(bytes).Pages[0];
        PdfTextSpan span = Assert.Single(page.GetTextSpans());
        Assert.Equal("gypqj", span.Text);
        Assert.True(span.ClipPath.HasValue);
        Assert.InRange(span.ClipPath.Value.Width, 26.99D, 27.01D);
        Assert.InRange(span.ClipPath.Value.Height, 8.99D, 9.01D);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        if (link) {
            var annotation = Assert.Single(independent.GetPage(1).GetHyperlinks());
            Assert.InRange(annotation.Bounds.Width, 0D, 27.01D);
            Assert.InRange(annotation.Bounds.Height, 0D, 9.01D);
        } else Assert.Empty(independent.GetPage(1).GetHyperlinks());
    }

    [Theory]
    [InlineData("direct", false)]
    [InlineData("clone", false)]
    [InlineData("nested", false)]
    [InlineData("direct", true)]
    [InlineData("clone", true)]
    [InlineData("nested", true)]
    public void SvgTextRetainsPaintBeyondItsLayoutFrame(string route, bool textLength) {
        string length = textLength ? " textLength='60' lengthAdjust='spacingAndGlyphs'" : "";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='200' height='80'>" +
            "<text x='20' y='40' font-family='Arial' font-size='20'" + length + ">ADWS</text></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing));
        Assert.NotNull(drawing);
        if (route == "clone") drawing = drawing.Clone();
        if (route == "nested") drawing = new OfficeDrawing(220, 100).AddDrawing(drawing, 10, 10);
        byte[] bytes = PdfDocument.Create().Drawing(drawing).ToBytes();

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetTextSpans());
        Assert.Equal("ADWS", span.Text);
        Assert.Null(span.ClipPath);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SvgAuthoredClipSurvivesCloneAndNestedDrawing(bool negativeOrigin, bool textLength) {
        string x = negativeOrigin ? "-2" : "20";
        string length = textLength ? " textLength='60' lengthAdjust='spacingAndGlyphs'" : "";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='200' height='80'>" +
            "<defs><clipPath id='cut'><rect width='35' height='60'/></clipPath></defs>" +
            "<text clip-path='url(#cut)' x='" + x + "' y='40' font-family='Arial' font-size='20'" + length + ">ADWS</text></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing));
        Assert.NotNull(drawing);
        drawing = new OfficeDrawing(220, 100).AddDrawing(drawing.Clone(), 10, 10);
        byte[] bytes = PdfDocument.Create().Drawing(drawing).ToBytes();

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetTextSpans());
        Assert.Equal("ADWS", span.Text);
        Assert.True(span.ClipPath.HasValue);
        Assert.InRange(span.ClipPath.Value.Width, 34.99D, 35.01D);
        Assert.InRange(span.ClipPath.Value.Height, 59.99D, 60.01D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PublicTextFramesKeepTheirExplicitClip(bool drawing) {
        byte[] bytes = drawing
            ? PdfDocument.Create().Drawing(new OfficeDrawing(100, 80)
                .AddPositionedText("ADWS", 10, 10, 20, 10, new OfficeFontInfo("Helvetica", 32), textAdvanceWidth: 80)).ToBytes()
            : PdfDocument.Create().Canvas(canvas => canvas.Text("ADWS", 10, 10, 20, 10, fontSize: 32)).ToBytes();

        var spans = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans().ToArray();
        Assert.NotEmpty(spans);
        Assert.All(spans, span => {
            Assert.True(span.ClipPath.HasValue);
            Assert.InRange(span.ClipPath.Value.Width, 19.99D, 20.01D);
            Assert.InRange(span.ClipPath.Value.Height, 9.99D, 10.01D);
        });
    }
}
