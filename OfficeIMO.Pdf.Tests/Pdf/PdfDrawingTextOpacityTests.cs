using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingTextOpacityTests {
    [Theory]
    [InlineData(0, false)]
    [InlineData(0, true)]
    [InlineData(128, false)]
    [InlineData(128, true)]
    public void DrawingTextRetainsAlphaWithoutChangingTheFollowingPaint(byte alpha, bool wrap) {
        var drawing = new OfficeDrawing(180D, 110D);
        drawing.AddText("FIRST", 0D, 0D, 180D, 40D, new OfficeFontInfo("Arial", 24D),
            OfficeColor.FromRgba(255, 0, 0, alpha), wrapText: wrap);
        drawing.AddText("FOLLOW", 0D, 60D, 180D, 40D, new OfficeFontInfo("Arial", 24D), OfficeColor.FromRgb(0, 0, 255));
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 220D, PageHeight = 150D,
            MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D
        });
        document.Content.Canvas(canvas => canvas.Drawing(drawing, 0D, 0D, 180D, 110D));
        byte[] bytes = document.ToBytes();
        string text = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("FIRST", text);
        Assert.Contains("FOLLOW", text);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes), 1D, OfficeColor.White);
        int firstInk = 0;
        int followingBlue = 0;
        for (int y = 20; y < 60; y++) {
            for (int x = 20; x < 200; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.R == 255 && pixel.G == 255 && pixel.B == 255) continue;
                firstInk++;
                Assert.Equal(255, pixel.R);
                Assert.True(pixel.G >= 126 && pixel.B >= 126, "Half-transparent red text must blend with the white page.");
            }
        }
        for (int y = 80; y < 120; y++)
            for (int x = 20; x < 200; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.B > 240 && pixel.R < 30 && pixel.G < 30) followingBlue++;
            }
        Assert.Equal(alpha == 0, firstInk == 0);
        Assert.True(followingBlue > 0, "The following opaque drawing text must retain its own paint state.");
    }
    [Theory]
    [InlineData(false, 128, 255)]
    [InlineData(true, 128, 255)]
    [InlineData(false, 255, 128)]
    [InlineData(true, 255, 128)]
    public void IndependentDecorationAlphaIsPreserved(bool wrap, byte textAlpha, byte decorationAlpha) {
        var drawing = new OfficeDrawing(180D, 110D);
        drawing.AddStyledText("FIRST", 0D, 0D, 180D, 40D, new OfficeFontInfo("Arial", 24D),
            OfficeColor.FromRgba(255, 0, 0, textAlpha), OfficeTextAlignment.Left, null,
            OfficeTextVerticalAlignment.Top, 0D, null, null, wrap, false, false, false, false,
            null, null, OfficeTextDecorationStyle.Single, OfficeTextDecorationStyle.Single,
            OfficeTextBaseline.Normal, 0, 1D, 0D, OfficeColor.FromRgba(0, 0, 255, decorationAlpha));
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 220D, PageHeight = 150D,
            MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D
        });
        document.Content.Canvas(canvas => canvas.Drawing(drawing, 0D, 0D, 180D, 110D));
        OfficeDrawing scene = PdfPageImageRenderer.RenderPage(document.ToBytes());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(scene, 2D, OfficeColor.White);
        int opaqueBlue = 0;
        int translucentBlue = 0;
        for (int y = 40; y < 120; y++)
            for (int x = 40; x < 400; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.B > 240 && pixel.R < 100 && pixel.G < 100) opaqueBlue++;
                if (pixel.B > 240 && pixel.R >= 100 && pixel.R < 230 && pixel.G >= 100 && pixel.G < 230) translucentBlue++;
            }
        Assert.Equal(decorationAlpha == 255, opaqueBlue > 0);
        if (decorationAlpha == 128) Assert.True(translucentBlue > 0, "Independent decoration paint must remain visible.");
    }
    [Theory]
    [InlineData(false, 128, 255)]
    [InlineData(true, 128, 255)]
    [InlineData(false, 255, 128)]
    [InlineData(true, 255, 128)]
    public void TaggedTransformedTextAppliesAlphaInsideIsolatedPaint(bool wrap, byte glyphAlpha, byte strokeAlpha) {
        var drawing = new OfficeDrawing(180D, 110D);
        drawing.AddStyledText("FIRST", 0D, 0D, 180D, 40D, new OfficeFontInfo("Arial", 24D),
            OfficeColor.FromRgba(255, 0, 0, glyphAlpha), OfficeTextAlignment.Left, null,
            OfficeTextVerticalAlignment.Top, 5D, null, null, wrap, false, false, false, false,
            null, null, OfficeTextDecorationStyle.Single, OfficeTextDecorationStyle.Single,
            OfficeTextBaseline.Normal, 0, 1D, 0D, OfficeColor.FromRgba(0, 0, 255, strokeAlpha));
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 220D, PageHeight = 150D, MarginLeft = 20D, MarginRight = 20D,
            MarginTop = 20D, MarginBottom = 20D, CompressContentStreams = false,
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
        });
        document.Content.Canvas(canvas => canvas.Drawing(drawing, 0D, 0D, 180D, 110D));
        string raw = System.Text.Encoding.GetEncoding(28591).GetString(document.ToBytes());
        var forms = System.Text.RegularExpressions.Regex.Matches(raw,
            @"(?s)/Subtype /Form(?<dictionary>.*?)stream\r?\n(?<content>.*?)\r?\nendstream");
        var form = Assert.Single(forms.Cast<System.Text.RegularExpressions.Match>());
        Assert.Contains("/I true", form.Groups["dictionary"].Value);
        // Isolated transparency Forms reset alpha; a caller gs cannot express separate
        // glyph/decorative alpha. Verify the state used by the inner paint, not our reader.
        var states = System.Text.RegularExpressions.Regex.Matches(form.Groups["content"].Value, @"/(?<name>GS\d+) gs");
        Assert.NotEmpty(states.Cast<System.Text.RegularExpressions.Match>());
        string name = states[0].Groups["name"].Value;
        var reference = System.Text.RegularExpressions.Regex.Match(form.Groups["dictionary"].Value,
            "/" + name + @" (?<id>\d+) 0 R");
        Assert.True(reference.Success);
        var state = System.Text.RegularExpressions.Regex.Match(raw,
            reference.Groups["id"].Value + @" 0 obj\s*<< /Type /ExtGState /ca (?<fill>[\d.]+) /CA (?<stroke>[\d.]+)");
        Assert.True(state.Success);
        Assert.Equal(glyphAlpha / 255D, double.Parse(state.Groups["fill"].Value, System.Globalization.CultureInfo.InvariantCulture), 3);
        Assert.Equal(strokeAlpha / 255D, double.Parse(state.Groups["stroke"].Value, System.Globalization.CultureInfo.InvariantCulture), 3);
    }

}
