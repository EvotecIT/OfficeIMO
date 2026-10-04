using System;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfDrawingNumericFontTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DrawingUsesTheRegisteredNumericFaceForLayoutAndPdfEmbedding(bool fitted) {
        byte[] regular = FontWithAdvance(500), medium = FontWithAdvance(1000);
        var fonts = new OfficeFontFaceCollection().Add("Scoped", regular, new OfficeFontFaceDescriptor(400))
            .Add("Scoped", medium, new OfficeFontFaceDescriptor(500));
        var options = new PdfOptions { PageWidth = 200, PageHeight = 100, MarginLeft = 0, MarginTop = 0, MarginRight = 0, MarginBottom = 0 };
        options.UseRenderingProfile(new OfficeRenderingProfile("numeric", fonts));
        var drawing = new OfficeDrawing(200, 100);
        var font = new OfficeFontInfo("Scoped", 20, new OfficeFontFaceDescriptor(500));
        drawing.AddText("A", 10, 10, 180, 70, font, wrapText: fitted);
        byte[] bytes = PdfDocument.Create(options).Compose(c => c.Page(p => p.Content(content => content.Drawing(drawing)))).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letter = Assert.Single(pdf.GetPage(1).Letters);
        double expected = OfficeTrueTypeFont.TryLoad(medium)!.Measure("A", 20);
        Assert.InRange(Math.Abs(letter.Width - expected), 0, .01D);
        Assert.Equal(font, Assert.Single(drawing.Elements.OfType<OfficeDrawingText>()).Font);
    }

    private static byte[] FontWithAdvance(int advance) {
        byte[] data = ManagedTextShapingTestAssets.CreateFont('A');
        int count = data[4] * 256 + data[5];
        for (int i = 0; i < count; i++) {
            int entry = 12 + i * 16;
            if (Encoding.ASCII.GetString(data, entry, 4) != "hmtx") continue;
            int offset = (data[entry + 8] << 24) | (data[entry + 9] << 16) | (data[entry + 10] << 8) | data[entry + 11];
            data[offset] = (byte)(advance >> 8); data[offset + 1] = (byte)advance;
            return data;
        }
        throw new InvalidOperationException("The fixture has no horizontal metrics table.");
    }
}
