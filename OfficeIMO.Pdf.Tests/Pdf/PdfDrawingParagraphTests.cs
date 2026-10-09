using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingParagraphTests {
    [Theory]
    [InlineData(OfficeTextAlignment.Center)]
    [InlineData(OfficeTextAlignment.Right)]
    public void UnwrappedOverwideParagraphRetainsItsPdfAlignmentAnchor(OfficeTextAlignment alignment) {
        const string caption = "UnwrappedOverflow";
        var drawing = new OfficeDrawing(200, 100).AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(caption, 10, OfficeColor.Black, fontFamily: "Arial") }, alignment)
        }, 40, 20, 40, 40, wrapText: false);
        byte[] bytes = PdfDocument.Create(new PdfOptions { PageWidth = 200, PageHeight = 100,
            MarginLeft = 0, MarginRight = 0, MarginTop = 0, MarginBottom = 0 })
            .Compose(canvas => canvas.Page(page => page.Content(content => content.Drawing(drawing)))).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(caption, string.Concat(letters.Select(letter => letter.Value)));
        double left = letters.Min(letter => letter.StartBaseLine.X), right = letters.Max(letter => letter.EndBaseLine.X);
        Assert.True(left < 40);
        if (alignment == OfficeTextAlignment.Center) Assert.InRange((left + right) / 2, 59.98, 60.02);
        else Assert.InRange(right, 79.98, 80.02);
    }

    [Fact]
    public void ParagraphOffsetsAndLeadingReachPdfAlongsideMixedRunStyles() {
        var drawing = new OfficeDrawing(200, 150).AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Left", 10, OfficeColor.Black) }, lineHeight: 20,
                margins: new OfficeTextPadding(10, 0, 10, 5)),
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Right", 10, OfficeColor.Red, bold: true) }, OfficeTextAlignment.Right,
                lineHeight: 30, margins: new OfficeTextPadding(10, 5, 15, 0))
        }, 20, 10, 160, 100);
        var bytes = PdfDocument.Create(new PdfOptions { PageWidth = 200, PageHeight = 150,
            MarginLeft = 0, MarginRight = 0, MarginTop = 0, MarginBottom = 0 })
            .Compose(c => c.Page(p => p.Content(content => content.Drawing(drawing)))).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters; Assert.Equal("LeftRight", string.Concat(letters.Select(l => l.Value)));
        var left = letters.Take(4).ToArray(); var right = letters.Skip(4).ToArray();
        Assert.InRange(left.Min(l => l.StartBaseLine.X), 29.99, 30.01);
        Assert.InRange(right.Max(l => l.EndBaseLine.X), 164.99, 165.01);
        Assert.InRange(left[0].StartBaseLine.Y - right[0].StartBaseLine.Y, 34.99, 35.01);
        Assert.All(right, l => Assert.Contains("Bold", l.FontName));
    }
}
