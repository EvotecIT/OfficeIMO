using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingLuminosityMaskWriterTests {
    [Fact]
    public void PdfLuminosityMaskPreservesCoefficientsPlacementAndGroupOpacity() {
        var content = new OfficeDrawing(20, 20)
            .AddShape(Rectangle(20, 20, OfficeColor.Blue), 0, 0)
            .AddShape(Rectangle(20, 20, OfficeColor.Blue), 0, 0);
        var mask = OfficeDrawingSoftMask.CreateWithLuminosityStandard(
            new OfficeDrawing(20, 20).AddShape(Rectangle(10, 20, OfficeColor.Red), 0, 0),
            OfficeSoftMaskLuminosityStandard.PdfDeviceRgb,
            OfficeSoftMaskMode.Luminosity, OfficeTransform.Translate(5, 0));
        var drawing = new OfficeDrawing(40, 40).AddEffectDrawing(content,
            OfficeTransform.Translate(10, 10), OfficeBlendMode.Normal, mask, 0.5);

        byte[] bytes = Write(drawing);
        string pdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/SMask << /S /Luminosity", pdf);
        Assert.Contains("/CS /DeviceRGB", pdf);
        Assert.Contains("/BC [0 0 0]", pdf);
        var raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes),
            background: OfficeColor.White);
        // PDF red luminosity is 0.3. Apply the completed mask, then group opacity:
        // white * (1 - 0.3 * 0.5) + blue * (0.3 * 0.5).
        OfficeColor inside = raster.GetPixel(20, 20);
        Assert.InRange(inside.R, (byte)216, (byte)218);
        Assert.InRange(inside.G, (byte)216, (byte)218);
        Assert.Equal((byte)255, inside.B);
        Assert.Equal(OfficeColor.White, raster.GetPixel(12, 20));
        Assert.Equal(OfficeColor.White, raster.GetPixel(28, 20));
    }

    [Fact]
    public void PrintConversionPreservesNestedLuminosityMaskVectorValues() {
        OfficeColor gray = OfficeColor.FromRgb(128, 128, 128);
        var content = new OfficeDrawing(20, 20).AddShape(Rectangle(20, 20, OfficeColor.Blue), 0, 0);
        var maskContent = new OfficeDrawing(20, 20).AddEffectDrawing(
            new OfficeDrawing(20, 20).AddShape(Rectangle(20, 20, gray), 0, 0),
            OfficeTransform.Identity, opacity: 0.5);
        var mask = OfficeDrawingSoftMask.CreateWithLuminosityStandard(maskContent,
            OfficeSoftMaskLuminosityStandard.PdfDeviceRgb, OfficeSoftMaskMode.Luminosity);
        var drawing = new OfficeDrawing(20, 20).AddEffectDrawing(content,
            OfficeTransform.Identity, OfficeBlendMode.Normal, mask);
        var options = PrintOptions(20, 20);

        byte[] bytes = Write(drawing, options);
        string pdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.502 0.502 0.502 rg", pdf);
        Assert.Contains(" k", pdf); // The visible blue still follows print conversion.
        var raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes));
        Assert.InRange(raster.GetPixel(10, 10).A, (byte)63, (byte)65);
    }

    [Fact]
    public void PrintConversionSeparatesMaskShadingsAndImagesFromVisibleResources() {
        var gradient = Rectangle(10, 10, OfficeColor.White);
        gradient.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        var pixel = new OfficeRasterImage(1, 1);
        pixel.SetPixel(0, 0, OfficeColor.FromRgb(128, 128, 128));
        byte[] image = OfficeRasterImageEncoder.Encode(pixel, OfficeImageExportFormat.Png);
        var shared = new OfficeDrawing(20, 10).AddShape(gradient, 0, 0)
            .AddImage(image, "image/png", new OfficeImageProjection(new OfficeImagePlacement(10, 0, 10, 10)));
        var mask = OfficeDrawingSoftMask.CreateWithLuminosityStandard(shared,
            OfficeSoftMaskLuminosityStandard.PdfDeviceRgb, OfficeSoftMaskMode.Luminosity);
        var drawing = new OfficeDrawing(20, 10).AddEffectDrawing(shared,
            OfficeTransform.Identity, OfficeBlendMode.Normal, mask);

        byte[] bytes = Write(drawing, PrintOptions(20, 10));
        var objects = PdfReadDocument.Open(bytes).Objects.Values.Select(item => item.Value).ToArray();
        var shadings = objects.OfType<PdfDictionary>().Where(dictionary =>
            dictionary.Items.ContainsKey("ShadingType")).ToArray();
        Assert.Single(shadings, dictionary => dictionary.Items["ColorSpace"] is PdfName { Name: "DeviceRGB" });
        Assert.Single(shadings, dictionary => dictionary.Items["ColorSpace"] is PdfName { Name: "DeviceCMYK" });
        var images = objects.OfType<PdfStream>().Select(stream => stream.Dictionary).Where(dictionary =>
            dictionary.Items.TryGetValue("Subtype", out var subtype) && subtype is PdfName { Name: "Image" }).ToArray();
        Assert.Single(images, dictionary => dictionary.Items["ColorSpace"] is PdfName { Name: "DeviceRGB" });
        Assert.Single(images, dictionary => dictionary.Items["ColorSpace"] is PdfName { Name: "DeviceCMYK" });
    }

    private static OfficeShape Rectangle(double width, double height, OfficeColor color) {
        var shape = OfficeShape.Rectangle(width, height);
        shape.FillColor = color;
        shape.StrokeWidth = 0;
        return shape;
    }

    private static PdfOptions PrintOptions(double width, double height) => Options(width, height)
        .ConfigurePdfXGroundwork(PdfComplianceProfile.PdfX4,
            IccMabTestProfiles.CreateCmykLab8Bidirectional(), "Managed test CMYK", PdfTrappingStatus.False);

    private static PdfOptions Options(double width, double height) => new PdfOptions {
        PageWidth = width, PageHeight = height, CompressContentStreams = false,
        MarginLeft = 0, MarginTop = 0, MarginRight = 0, MarginBottom = 0
    };

    private static byte[] Write(OfficeDrawing drawing, PdfOptions? options = null) =>
        PdfDocument.Create(options ?? Options(drawing.Width, drawing.Height)).Drawing(drawing).ToBytes();
}
