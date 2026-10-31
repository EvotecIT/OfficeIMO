using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointNativeThemeTintTests {
    [Fact]
    public void TintShadeOrderAndAlphaMatchSevenNativePowerPointControls() {
        string folder = Path.Combine(AppContext.BaseDirectory, "Documents", "PowerPointNativeThemeTint");
        using var presentation = PowerPointDocument.Load(Path.Combine(folder, "controls.pptx"));
        PowerPointSlide slide = Assert.Single(presentation.Slides);
        PowerPointAutoShape[] controls = slide.Shapes.OfType<PowerPointAutoShape>().ToArray();
        Assert.Equal(7, controls.Length);
        Assert.Equal(OfficeColor.FromRgba(204, 210, 216, 166), controls[6].ResolveFillColorForExport());
        OfficeRasterImage nativePng = VisualBaselineTestSupport.DecodePng(
            File.ReadAllBytes(Path.Combine(folder, "controls-native.png")), "Native color controls are not a valid PNG.");
        byte[] reference = File.ReadAllBytes(Path.Combine(folder, "microsoft-powerpoint-16.113.4-controls.pdf"));
        OfficeRasterImage nativePdf = OfficeDrawingRasterRenderer.Render(PdfDocument.Load(new System.IO.MemoryStream(reference)).Pages[0].ToDrawing(),
            background: OfficeColor.White);
        OfficeRasterImage managed = OfficeDrawingRasterRenderer.Render(PdfDocument.Load(new System.IO.MemoryStream(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).Pages[0].ToDrawing(),
            background: OfficeColor.White);
        OfficeColor[] expected = {
            OfficeColor.FromRgb(21, 96, 130), OfficeColor.FromRgb(204, 210, 216),
            OfficeColor.FromRgb(170, 170, 170), OfficeColor.FromRgb(250, 226, 219),
            OfficeColor.FromRgb(135, 139, 143), OfficeColor.FromRgb(204, 206, 209),
            OfficeColor.FromRgb(221, 225, 229)
        };

        AssertControlSamples(nativePng, nativePdf, managed, expected, firstAlphaIndex: 6);
    }

    [Fact]
    public void LuminanceOrderAndInterleavedAlphaMatchSevenNativePowerPointControls() {
        string folder = Path.Combine(AppContext.BaseDirectory, "Documents", "PowerPointNativeThemeTint");
        using var presentation = PowerPointDocument.Load(Path.Combine(folder, "luminance-controls.pptx"));
        PowerPointSlide slide = Assert.Single(presentation.Slides);
        PowerPointAutoShape[] controls = slide.Shapes.OfType<PowerPointAutoShape>().ToArray();
        Assert.Equal(7, controls.Length);
        // Alpha between HSL operations must not force an intermediate encoded-RGB rounding step.
        OfficeColor translucent = OfficeColor.FromRgba(56, 100, 178, 115);
        Assert.Equal(translucent, controls[5].ResolveFillColorForExport());
        Assert.Equal(translucent, controls[6].ResolveFillColorForExport());
        OfficeRasterImage nativePng = VisualBaselineTestSupport.DecodePng(
            File.ReadAllBytes(Path.Combine(folder, "luminance-controls-native.png")), "Native luminance controls are not a valid PNG.");
        byte[] reference = File.ReadAllBytes(Path.Combine(folder, "microsoft-powerpoint-16.113.4-luminance-controls.pdf"));
        OfficeRasterImage nativePdf = OfficeDrawingRasterRenderer.Render(PdfDocument.Load(new System.IO.MemoryStream(reference)).Pages[0].ToDrawing(),
            background: OfficeColor.White);
        OfficeRasterImage managed = OfficeDrawingRasterRenderer.Render(PdfDocument.Load(new System.IO.MemoryStream(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).Pages[0].ToDrawing(),
            background: OfficeColor.White);
        OfficeColor[] expected = {
            OfficeColor.FromRgb(68, 114, 196), OfficeColor.FromRgb(32, 56, 100),
            OfficeColor.FromRgb(146, 172, 220), OfficeColor.FromRgb(56, 100, 178),
            OfficeColor.FromRgb(44, 78, 139), OfficeColor.FromRgb(165, 185, 220),
            OfficeColor.FromRgb(165, 185, 220)
        };

        AssertControlSamples(nativePng, nativePdf, managed, expected, firstAlphaIndex: 5);
    }

    [Fact]
    public void TableBandTintMatchesNativePowerPointLinearRgbSample() {
        string folder = Path.Combine(AppContext.BaseDirectory, "Documents", "PowerPointNativeThemeTint");
        using var presentation = PowerPointDocument.Load(Path.Combine(folder, "p1-tables-charts.pptx"));
        PowerPointTable table = Assert.Single(presentation.Slides[1].Tables);
        OfficeColor expected = OfficeColor.FromRgb(204, 210, 216);

        // Medium Style 2 - Accent 1 uses accent 156082 and tint=40000 for band1H.
        Assert.Equal(OfficeColor.FromRgb(21, 96, 130), PowerPointSlideImageRenderer.ResolveTableCellFillColorForExport(table, 0, 0));
        Assert.Equal(expected, PowerPointSlideImageRenderer.ResolveTableCellFillColorForExport(table, 1, 0));

        byte[] reference = File.ReadAllBytes(Path.Combine(folder, "microsoft-powerpoint-16.113.4-tables-charts.pdf"));
        OfficeRasterImage native = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(reference).Pages[1].ToDrawing(),
            background: OfficeColor.White);
        OfficeRasterImage managed = OfficeDrawingRasterRenderer.Render(PdfDocument.Load(new System.IO.MemoryStream(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).Pages[1].ToDrawing(),
            background: OfficeColor.White);
        // A clear point inside the second row, away from text and borders, in PDF page coordinates.
        Assert.Equal(expected, native.GetPixel(100, 185));
        Assert.Equal(expected, managed.GetPixel(100, 185));
    }

    private static void AssertControlSamples(OfficeRasterImage nativePng, OfficeRasterImage nativePdf,
        OfficeRasterImage managed, OfficeColor[] expected, int firstAlphaIndex) {
        for (int index = 0; index < expected.Length; index++) {
            int x = 35 + index * 65;
            Assert.Equal(expected[index], nativePng.GetPixel(x, 70));
            // Only alpha controls permit one channel value for rounded eight-bit alpha compositing.
            int tolerance = index >= firstAlphaIndex ? 1 : 0;
            foreach ((string route, OfficeRasterImage image) in new[] { ("native PDF", nativePdf), ("managed PDF", managed) }) {
                OfficeColor pixel = image.GetPixel(x, 70);
                Assert.True(Math.Abs(pixel.R - expected[index].R) <= tolerance &&
                    Math.Abs(pixel.G - expected[index].G) <= tolerance && Math.Abs(pixel.B - expected[index].B) <= tolerance,
                    $"Control {index} in {route}: expected {expected[index]}, observed {pixel}.");
                Assert.Equal(255, pixel.A);
            }
        }
    }
}
