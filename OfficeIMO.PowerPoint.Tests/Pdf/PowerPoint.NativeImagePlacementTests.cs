using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PowerPointSaveAsPdfTests {
    [Fact]
    public void SaveAsPdf_UnsupportedSvgPictureReportsOmissionWithoutAbortingTheSlide() {
        using var presentation = PowerPointPresentation.Create();
        presentation.SlideSize.SetSizePoints(160, 100);
        PowerPointSlide slide = presentation.AddSlide();
        slide.BackgroundColor = "FFFFFF";
        byte[] svg = System.Text.Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>"
            + "<foreignObject width='10' height='10'><div xmlns='http://www.w3.org/1999/xhtml'>Visible HTML</div></foreignObject></svg>");
        slide.AddPicture(new MemoryStream(svg), OfficeImageFormat.Svg, 0, 0,
            PowerPointUnits.FromPoints(40), PowerPointUnits.FromPoints(30));
        slide.AddShapePoints(OfficePresetShapeType.Rectangle, 80, 20, 30, 20).FillColor = "D7263D";

        PdfDocumentConversionResult result = presentation.ToPdfDocumentResult();
        byte[] pdf = result.ToBytes();

        Assert.Contains(result.Warnings, warning => warning.Code == "unsupported-picture");
        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(),
            background: OfficeColor.White);
        Assert.Equal(OfficeColor.FromRgb(0xD7, 0x26, 0x3D), rendered.GetPixel(95, 30));
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RotatesPictureCropWithImageFrame() {
        string folder = Path.Combine(AppContext.BaseDirectory, "Documents", "PowerPointNativeImagePlacement");
        using var presentation = PowerPointPresentation.Load(Path.Combine(folder, "picture-crop-rotation.pptx"));
        byte[] actual = presentation.ToPdfBytes();
        byte[] reference = File.ReadAllBytes(Path.Combine(folder, "microsoft-powerpoint-16.113.4-picture-crop-rotation.pdf"));
        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(actual).Pages[0].ToDrawing(), 1D, OfficeColor.White);
        OfficeRasterImage native = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(reference).Pages[0].ToDrawing(), 1D, OfficeColor.White);

        // PowerPoint rotates the cropped viewport clockwise about the authored
        // frame, not about the expanded source image. The asymmetric colors also
        // distinguish the rotation direction; the old solid-color matrix did not.
        foreach ((int x, int y) in new[] { (70, 40), (70, 55), (70, 80), (70, 90), (50, 65), (90, 65), (70, 30), (70, 100) }) {
            OfficeColor expected = native.GetPixel(x, y);
            OfficeColor pixel = rendered.GetPixel(x, y);
            Assert.InRange(Math.Abs(expected.R - pixel.R), 0, 3);
            Assert.InRange(Math.Abs(expected.G - pixel.G), 0, 3);
            Assert.InRange(Math.Abs(expected.B - pixel.B), 0, 3);
            Assert.Equal(expected.A, pixel.A);
        }
    }
}
