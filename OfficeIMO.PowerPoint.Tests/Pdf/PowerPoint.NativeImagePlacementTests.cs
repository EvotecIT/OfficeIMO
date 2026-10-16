using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PowerPointSaveAsPdfTests {
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
