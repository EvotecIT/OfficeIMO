using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRendererFidelityTests {
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(5)]
    [InlineData(6)]
    public void GlyphOutlineStrokesRetainTheirDashPatternAndPhase(int mode) {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        var objects = new List<(string, byte[])> {
            ("<< /Length " + font.Length + " >>", font),
            ("<< /Type /FontDescriptor /FontName /Fixture /Flags 32 /FontBBox [0 0 1000 1000] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 /FontFile2 5 0 R >>", Array.Empty<byte>())
        };
        const string resources = "<< /Font << /F1 << /Type /Font /Subtype /TrueType /BaseFont /Fixture /Encoding /WinAnsiEncoding /FirstChar 65 /LastChar 65 /Widths [600] /FontDescriptor 6 0 R >> >> >>";
        byte[] Document(string dash) => BuildPdf("0 0 1 rg 1 0 0 RG 2 w " + dash + " d BT /F1 48 Tf " + mode + " Tr 20 25 Td (A) Tj ET", resources, objects: objects);
        OfficeRasterImage solid = Render(Document("[] 0"));
        OfficeRasterImage dashed = Render(Document("[3 3] 1"));
        int Red(OfficeRasterImage image) => CountPixels(image, 0, 0, image.Width, image.Height,
            color => color.A > 100 && color.R > 180 && color.G < 100 && color.B < 100);
        Assert.True(Red(solid) > 20);
        Assert.InRange(Red(dashed), 10, (int)(Red(solid) * .85D));
        OfficeDrawing drawing = PdfReadDocument.Open(Document("[3 3] 1")).Pages[0].ToDrawing();
        var options = new OfficeDrawingRasterRenderOptions { PdfRasterConventions = true };
        OfficeRasterImage direct = OfficeDrawingRasterRenderer.Render(drawing, options);
        OfficeRasterImage clone = OfficeDrawingRasterRenderer.Render(drawing.Clone(), options);
        Assert.Equal(direct.GetPixels(), clone.GetPixels());
    }
}
