using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public sealed class PdfDrawingLuminosityMaskWriterTests {
        [Fact]
        public void PdfLuminosityMaskPreservesCoefficientsPlacementAndGroupOpacity() {
            OfficeDrawing content = new OfficeDrawing(20, 20)
                .AddShape(Rectangle(20, 20, OfficeColor.Blue), 0, 0)
                .AddShape(Rectangle(20, 20, OfficeColor.Blue), 0, 0);
            OfficeDrawingSoftMask mask = OfficeDrawingSoftMask.CreateWithLuminosityStandard(
                new OfficeDrawing(20, 20).AddShape(Rectangle(10, 20, OfficeColor.Red), 0, 0),
                OfficeSoftMaskLuminosityStandard.PdfDeviceRgb,
                OfficeSoftMaskMode.Luminosity, OfficeTransform.Translate(5, 0));
            OfficeDrawing drawing = new OfficeDrawing(40, 40).AddEffectDrawing(content,
                OfficeTransform.Translate(10, 10), OfficeBlendMode.Normal, mask, 0.5);

            byte[] bytes = Write(drawing);
            string pdf = Encoding.ASCII.GetString(bytes);
            Assert.Contains("/SMask << /S /Luminosity", pdf);
            Assert.Contains("/CS /DeviceRGB", pdf);
            Assert.Contains("/BC [0 0 0]", pdf);
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes),
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

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PrintConversionPreservesNestedLuminosityMaskVectorValues(bool useTilingPattern) {
            OfficeColor gray = OfficeColor.FromRgb(128, 128, 128);
            OfficeDrawing content = new OfficeDrawing(20, 20).AddShape(Rectangle(20, 20, OfficeColor.Blue), 0, 0);
            OfficeDrawing tile = new OfficeDrawing(20, 20).AddShape(Rectangle(20, 20, gray), 0, 0);
            OfficeDrawing maskContent = useTilingPattern
                ? new OfficeDrawing(20, 20).AddTilingPattern(tile, new OfficeImagePlacement(0, 0, 20, 20), 20, 20, opacity: 0.5)
                : new OfficeDrawing(20, 20).AddEffectDrawing(tile, OfficeTransform.Identity, opacity: 0.5);
            OfficeDrawingSoftMask mask = OfficeDrawingSoftMask.CreateWithLuminosityStandard(maskContent,
                OfficeSoftMaskLuminosityStandard.PdfDeviceRgb, OfficeSoftMaskMode.Luminosity);
            OfficeDrawing drawing = new OfficeDrawing(20, 20).AddEffectDrawing(content,
                OfficeTransform.Identity, OfficeBlendMode.Normal, mask);
            PdfCreateOptions options = PrintOptions(20, 20);

            byte[] bytes = Write(drawing, options);
            PdfStream[] groups = PdfReadDocument.Open(bytes).Objects.Values.Select(item => item.Value)
                .OfType<PdfStream>().Where(stream => stream.Dictionary.Items.ContainsKey("Group")).ToArray();
            Assert.Contains(groups, stream => stream.Dictionary.Items["Group"] is PdfDictionary group &&
                group.Items.TryGetValue("CS", out PdfObject? cs) && cs is PdfName { Name: "DeviceRGB" } &&
                (!useTilingPattern || stream.Dictionary.Items["Resources"] is PdfDictionary resources &&
                    resources.Items.ContainsKey("Pattern")));
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes));
            Assert.InRange(raster.GetPixel(10, 10).A, (byte)63, (byte)65);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PrintConversionSeparatesMaskShadingsAndImagesFromVisibleResources(bool radial) {
            OfficeShape gradient = Rectangle(10, 10, OfficeColor.White);
            OfficeColor transparentBlue = OfficeColor.FromRgba(0, 0, 255, 128);
            OfficeShape separateAlpha = Rectangle(10, 10, OfficeColor.White);
            if (radial) {
                gradient.FillRadialGradient = OfficeRadialGradient.Centered(OfficeColor.Red, transparentBlue);
                separateAlpha.FillRadialGradient = gradient.FillRadialGradient.WithSeparateAlphaInterpolation();
            } else {
                gradient.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, transparentBlue);
                separateAlpha.FillGradient = gradient.FillGradient.WithSeparateAlphaInterpolation();
            }
            OfficeRasterImage pixel = new OfficeRasterImage(1, 1);
            pixel.SetPixel(0, 0, OfficeColor.FromRgb(64, 128, 192));
            byte[] image = OfficeRasterImageEncoder.Encode(pixel, OfficeImageExportFormat.Png);
            OfficeDrawing shared = new OfficeDrawing(20, 10).AddShape(gradient, 0, 0).AddShape(separateAlpha, 0, 0)
                .AddImage(image, "image/png", new OfficeImageProjection(new OfficeImagePlacement(10, 0, 10, 10)));
            OfficeDrawingSoftMask mask = OfficeDrawingSoftMask.CreateWithLuminosityStandard(shared,
                OfficeSoftMaskLuminosityStandard.PdfDeviceRgb, OfficeSoftMaskMode.Luminosity);
            OfficeDrawing drawing = new OfficeDrawing(20, 10).AddEffectDrawing(shared,
                OfficeTransform.Identity, OfficeBlendMode.Normal, mask);

            byte[] bytes = Write(drawing, PrintOptions(20, 10));
            PdfObject[] objects = PdfReadDocument.Open(bytes).Objects.Values.Select(item => item.Value).ToArray();
            PdfDictionary[] shadings = objects.OfType<PdfDictionary>().Where(dictionary =>
                dictionary.Items.ContainsKey("ShadingType")).ToArray();
            Assert.Equal(2, shadings.Count(dictionary => dictionary.Items["ColorSpace"] is PdfName { Name: "DeviceRGB" }));
            Assert.Equal(2, shadings.Count(dictionary => dictionary.Items["ColorSpace"] is PdfName { Name: "DeviceCMYK" }));
            PdfDictionary[] images = objects.OfType<PdfStream>().Select(stream => stream.Dictionary).Where(dictionary =>
                dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) && subtype is PdfName { Name: "Image" }).ToArray();
            Assert.Single(images, dictionary => dictionary.Items["ColorSpace"] is PdfName { Name: "DeviceRGB" });
            Assert.Single(images, dictionary => dictionary.Items["ColorSpace"] is PdfName { Name: "DeviceCMYK" });
        }

        private static OfficeShape Rectangle(double width, double height, OfficeColor color) {
            OfficeShape shape = OfficeShape.Rectangle(width, height);
            shape.FillColor = color;
            shape.StrokeWidth = 0;
            return shape;
        }

        private static PdfCreateOptions PrintOptions(double width, double height) => Options(width, height)
            .ConfigurePdfXGroundwork(PdfComplianceProfile.PdfX4,
                IccMabTestProfiles.CreateCmykLab8Bidirectional(), "Managed test CMYK", PdfTrappingStatus.False);

        private static PdfCreateOptions Options(double width, double height) => new PdfCreateOptions { BackgroundColor = null,
            PageWidth = width, PageHeight = height, CompressContentStreams = false,
            MarginLeft = 0, MarginTop = 0, MarginRight = 0, MarginBottom = 0
        };

        private static byte[] Write(OfficeDrawing drawing, PdfCreateOptions? options = null) =>
            PdfDocument.Create(options ?? Options(drawing.Width, drawing.Height)).Drawing(drawing).ToBytes();
    }
}
