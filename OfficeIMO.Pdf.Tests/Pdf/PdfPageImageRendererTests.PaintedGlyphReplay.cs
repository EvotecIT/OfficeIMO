using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData(false, "fi", "00660069")]
    [InlineData(true, "fi", "00660069")]
    [InlineData(false, "ب", "0628")]
    [InlineData(true, "ب", "0628")]
    public void LaterTransparencyFormKeepsPaintedGlyphMappings(bool nested, string logicalText, string unicode) {
        byte[] pdf = BuildPaintedGlyphReplayPdf(nested, unicode);
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        OfficeDrawing drawing = page.ToDrawing();
        var text = ReplayText(drawing).ToArray();

        Assert.Equal(2, text.Length);
        Assert.All(text, item => {
            Assert.Equal(logicalText, item.Text.Text);
            Assert.NotEqual(logicalText, item.Text.RasterText);
            Assert.True(item.Owner.Fonts.TryResolveFaceForText(item.Text.RasterText, item.Text.Font.FamilyName,
                item.Text.Font.Style, out OfficeFontFace? face), "The retained drawing must resolve its exact painted glyph.");
            Assert.True(face!.Program.TryGetGlyphMetrics(item.Text.RasterText[0], out int glyph, out _));
            Assert.Equal(1, glyph);
        });
        Assert.Equal(logicalText + logicalText, page.ExtractText().Replace(" ", string.Empty).Replace("\n", string.Empty));

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
        // The embedded glyph is a known 400-by-700 rectangle, not a substitute for its logical letters.
        Assert.Equal(OfficeColor.Black, raster.GetPixel(23, 90));
        Assert.Equal(OfficeColor.White, raster.GetPixel(29, 90));
        Assert.Equal(OfficeColor.Black, raster.GetPixel(83, 90));
        Assert.Equal(OfficeColor.White, raster.GetPixel(89, 90));
    }

    private static IEnumerable<(OfficeDrawing Owner, OfficeDrawingText Text)> ReplayText(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingText text) yield return (drawing, text);
            OfficeDrawing? inner = element switch {
                OfficeDrawingGroup group => group.InnerDrawing,
                OfficeDrawingEffectGroup group => group.InnerDrawing,
                _ => null
            };
            if (inner != null) foreach (var child in ReplayText(inner)) yield return child;
        }
    }

    private static byte[] BuildPaintedGlyphReplayPdf(bool nested, string unicode) {
        const string painted = "BT /F1 20 Tf 1 0 0 1 20 100 Tm (A) Tj ET /Reset Do BT /F1 20 Tf 1 0 0 1 80 100 Tm (A) Tj ET";
        const string resources = "<< /Font << /F1 5 0 R >> /XObject << /Reset 10 0 R /Outer 11 0 R >> >>";
        const string outerResources = "<< /Font << /F1 5 0 R >> /XObject << /Reset 10 0 R >> >>";
        return BuildSingleStreamPdf(nested ? "/Outer Do" : painted, resources,
            "5 0 obj\n<< /Type /Font /Subtype /TrueType /BaseFont /Replay /FirstChar 65 /LastChar 65 /Widths [500] /FontDescriptor 7 0 R /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding /Differences [65 /A] >> /ToUnicode 9 0 R >>\nendobj",
            "7 0 obj\n<< /Type /FontDescriptor /FontName /Replay /Flags 32 /FontBBox [0 0 600 700] /Ascent 800 /Descent -200 /CapHeight 700 /ItalicAngle 0 /StemV 80 /FontFile2 8 0 R >>\nendobj",
            BuildStreamObject(8, "<< /Filter /ASCIIHexDecode", BitConverter.ToString(
                ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphOutlines('A', 'B')).Replace("-", string.Empty) + ">"),
            BuildStreamObject(9, "<<", "begincmap\n1 begincodespacerange\n<00> <FF>\nendcodespacerange\n1 beginbfchar\n<41> <" + unicode + ">\nendbfchar\nendcmap"),
            BuildStreamObject(10, "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Group << /S /Transparency /I true /K false >> /Resources << >>",
                "0 1 0 rg 180 20 10 10 re f"),
            BuildStreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Group << /S /Transparency /I true /K false >> /Resources " + outerResources,
                painted));
    }
}
