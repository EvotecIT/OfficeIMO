using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingPositionedTextTests {
    [Fact]
    public void SinglePositionedLineKeepsItsExplicitAdvance() {
        var drawing = new OfficeDrawing(100, 80).AddPositionedText("A", 10, 10, 80, 60,
            new OfficeFontInfo("Helvetica", 20), alignment: OfficeTextAlignment.Center, textAdvanceWidth: 50);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
            PageWidth = 100, PageHeight = 80, MarginLeft = 0, MarginTop = 0, MarginRight = 0, MarginBottom = 0
        }).Compose(c => c.Page(p => p.Content(content => content.Drawing(drawing)))).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letter = Assert.Single(pdf.GetPage(1).Letters);
        Assert.InRange(letter.Width, 49.99D, 50.01D);
        Assert.InRange(letter.Location.X, 24.99D, 25.01D);
    }

    [Theory]
    [InlineData(OfficeTextAlignment.Center, "\n")]
    [InlineData(OfficeTextAlignment.Right, "\r\n")]
    [InlineData(OfficeTextAlignment.Right, "\r")]
    public void PositionedHardLinesRetainTheirOwnAdvance(OfficeTextAlignment alignment, string separator) {
        var drawing = new OfficeDrawing(100, 80).AddPositionedText("AAAA" + separator + "A", 10, 10, 80, 60,
            new OfficeFontInfo("Helvetica", 20), alignment: alignment, lineHeight: 25);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
            PageWidth = 100, PageHeight = 80, MarginLeft = 0, MarginTop = 0, MarginRight = 0, MarginBottom = 0
        }).Compose(c => c.Page(p => p.Content(content => content.Drawing(drawing)))).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal("AAAAA", string.Concat(letters.Select(letter => letter.Value)));
        Assert.InRange(Math.Abs(letters[0].Width - letters[4].Width), 0, .01D);
        Assert.InRange(letters[0].Width, 13.3D, 13.4D);
        double firstAdvance = letters[0].Width * 4;
        double lastAdvance = letters[4].Width;
        double factor = alignment == OfficeTextAlignment.Center ? .5D : 1D;
        Assert.InRange(Math.Abs(letters[0].Location.X - (10 + (80 - firstAdvance) * factor)), 0, .01D);
        Assert.InRange(Math.Abs(letters[4].Location.X - (10 + (80 - lastAdvance) * factor)), 0, .01D);
        Assert.InRange(Math.Abs(letters[0].Location.Y - letters[4].Location.Y - 25), 0, .01D);
    }
}
