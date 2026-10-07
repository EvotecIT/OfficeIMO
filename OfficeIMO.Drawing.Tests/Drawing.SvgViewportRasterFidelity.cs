using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class DrawingSvgViewportRasterFidelityTests {
        private const string Bars = "<svg xmlns='http://www.w3.org/2000/svg' width='60' height='24' viewBox='0 0 20 8'>"
            + "<rect width='20' height='8' fill='white'/>"
            + "<g fill='black'><rect x='2' width='2' height='6'/><rect x='5' width='1' height='6'/>"
            + "<rect x='7' width='1' height='6'/><rect x='10' width='3' height='6'/></g></svg>";

        [Theory]
        [InlineData(1)]
        [InlineData(2)]
        public void ScaledSvgViewportPreservesIntegerAlignedBarEdges(int outputScale) {
            OfficeDrawing drawing = ReadBars();

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, outputScale);

            AssertBars(image, outputScale);
        }

        [Fact]
        public void PointConversionAndDpiScalingPreserveOriginalSvgBarEdges() {
            OfficeDrawing pixels = ReadBars();
            OfficeDrawing points = new OfficeDrawing(pixels.Width * 0.75D, pixels.Height * 0.75D);
            points.AddEffectDrawing(pixels, OfficeTransform.Scale(0.75D, 0.75D));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(points, 96D / 72D);

            AssertBars(image, 1);
        }

        [Theory]
        [InlineData(0.5D)]
        [InlineData(1.25D)]
        public void FractionalEffectScaleMatchesDirectVectorCoverage(double groupScale) {
            OfficeDrawing svg = ReadBars();
            OfficeDrawing transformed = new OfficeDrawing(svg.Width * groupScale, svg.Height * groupScale);
            transformed.AddEffectDrawing(svg, OfficeTransform.Scale(groupScale, groupScale));
            OfficeDrawing reference = new OfficeDrawing(20D, 8D);
            AddRectangle(reference, 0D, 20D, 8D, OfficeColor.White);
            foreach ((double x, double width) in new[] { (2D, 2D), (5D, 1D), (7D, 1D), (10D, 3D) }) {
                AddRectangle(reference, x, width, 6D, OfficeColor.Black);
            }

            OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(transformed);
            OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(reference, 3D * groupScale);

            Assert.Equal(expected.Width, actual.Width);
            Assert.Equal(expected.Height, actual.Height);
            Assert.Equal(expected.GetPixels(), actual.GetPixels());
        }

        [Fact]
        public void MagnifiedEffectLayerHonorsIntermediatePixelBudget() {
            OfficeDrawing inner = new OfficeDrawing(100D, 100D);
            OfficeShape rectangle = OfficeShape.Rectangle(100D, 100D);
            rectangle.FillColor = OfficeColor.Black;
            rectangle.StrokeColor = null;
            inner.AddShape(rectangle, 0D, 0D);
            OfficeDrawing drawing = new OfficeDrawing(100D, 100D);
            drawing.AddEffectDrawing(inner, OfficeTransform.Scale(8D, 8D));

            Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 10_000 }));
        }

        [Theory]
        [InlineData(OfficeBlendMode.Normal, false)]
        [InlineData(OfficeBlendMode.Multiply, false)]
        [InlineData(OfficeBlendMode.Normal, true)]
        [InlineData(OfficeBlendMode.Multiply, true)]
        public void FractionalNonuniformScalePreservesNearestImageGrid(OfficeBlendMode blend, bool coloredTexels) {
            OfficeRasterImage source = new OfficeRasterImage(3, 3, OfficeColor.Black);
            if (coloredTexels) {
                for (int y = 0; y < 3; y++) for (int x = 0; x < 3; x++) {
                    source.SetPixel(x, y, OfficeColor.FromRgb((byte)(x * 100), (byte)(y * 100), (byte)((x + y) * 50)));
                }
            }
            byte[] png = OfficePngWriter.Encode(source);
            OfficeDrawing inner = new OfficeDrawing(3D, 3D);
            inner.AddImageWithInterpolation(png, "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 3D, 3D)), interpolate: false);
            OfficeDrawing grouped = new OfficeDrawing(4D, 3D);
            grouped.AddEffectDrawing(inner, OfficeTransform.Scale(1.1D, 0.9D), blend);
            OfficeDrawing direct = new OfficeDrawing(4D, 3D);
            direct.AddImageWithInterpolation(png, "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 3.3D, 2.7D)), interpolate: false);

            OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(grouped);
            OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(direct);

            Assert.Equal(expected.GetPixels(), actual.GetPixels());
            Assert.Equal((byte)255, actual.GetPixel(1, 2).A);
        }

        [Theory]
        [InlineData(OfficeBlendMode.Normal)]
        [InlineData(OfficeBlendMode.Multiply)]
        public void FractionalNonuniformScalePreservesNearestSoftMaskGrid(OfficeBlendMode blend) {
            OfficeRasterImage maskPixels = new OfficeRasterImage(3, 3, OfficeColor.Black);
            OfficeRasterImage expectedPixels = new OfficeRasterImage(3, 3, OfficeColor.Transparent);
            for (int y = 0; y < 3; y++) for (int x = 0; x < 3; x++) {
                if ((x + y) % 2 == 0) {
                    maskPixels.SetPixel(x, y, OfficeColor.White);
                    expectedPixels.SetPixel(x, y, OfficeColor.Red);
                }
            }
            OfficeDrawing maskDrawing = new OfficeDrawing(3D, 3D);
            maskDrawing.AddImageWithInterpolation(OfficePngWriter.Encode(maskPixels), "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 3D, 3D)), interpolate: false);
            OfficeDrawing source = new OfficeDrawing(3D, 3D);
            AddRectangle(source, 0D, 3D, 3D, OfficeColor.Red);
            OfficeDrawing grouped = new OfficeDrawing(4D, 3D);
            grouped.AddEffectDrawing(source, OfficeTransform.Scale(1.1D, 0.9D), blend,
                new OfficeDrawingSoftMask(maskDrawing, OfficeSoftMaskMode.Luminosity));
            OfficeDrawing direct = new OfficeDrawing(4D, 3D);
            direct.AddImageWithInterpolation(OfficePngWriter.Encode(expectedPixels), "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 3.3D, 2.7D)), interpolate: false);

            Assert.Equal(OfficeDrawingRasterRenderer.Render(direct).GetPixels(),
                OfficeDrawingRasterRenderer.Render(grouped).GetPixels());
        }

        private static OfficeDrawing ReadBars() {
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(Bars), out OfficeDrawing? drawing, out int unsupported));
            Assert.Equal(0, unsupported);
            return drawing!;
        }

        private static void AddRectangle(OfficeDrawing drawing, double x, double width, double height, OfficeColor color) {
            OfficeShape rectangle = OfficeShape.Rectangle(width, height);
            rectangle.FillColor = color;
            rectangle.StrokeColor = null;
            drawing.AddShape(rectangle, x, 0D);
        }

        private static void AssertBars(OfficeRasterImage image, int outputScale) {
            Assert.Equal(60 * outputScale, image.Width);
            Assert.Equal(24 * outputScale, image.Height);
            for (int x = 0; x < image.Width; x++) {
                int module = x / (3 * outputScale);
                bool dark = module is 2 or 3 or 5 or 7 or 10 or 11 or 12;
                Assert.Equal(dark ? OfficeColor.Black : OfficeColor.White, image.GetPixel(x, 6 * outputScale));
            }
        }
    }

    public partial class DrawingTests {
        [Theory]
        [InlineData(OfficeBlendMode.Normal, false)]
        [InlineData(OfficeBlendMode.Multiply, false)]
        [InlineData(OfficeBlendMode.Normal, true)]
        [InlineData(OfficeBlendMode.Multiply, true)]
        public void OfficeDrawingEffectGroup_ExhaustedInspectionPreservesLaterNearestImageAndMaskGrid(OfficeBlendMode blend, bool useMask) {
            OfficeRasterImage pixels = new OfficeRasterImage(3, 3);
            OfficeRasterImage expectedPixels = new OfficeRasterImage(3, 3);
            for (int y = 0; y < 3; y++) for (int x = 0; x < 3; x++) {
                if (useMask) {
                    bool opaque = (x + y) % 2 == 0;
                    pixels.SetPixel(x, y, opaque ? OfficeColor.White : OfficeColor.Black);
                    expectedPixels.SetPixel(x, y, opaque ? OfficeColor.Red : OfficeColor.Transparent);
                } else {
                    OfficeColor color = OfficeColor.FromRgb((byte)(x * 100), (byte)(y * 100), (byte)((x + y) * 50));
                    pixels.SetPixel(x, y, color);
                    expectedPixels.SetPixel(x, y, color);
                }
            }
            byte[] png = OfficePngWriter.Encode(pixels);
            OfficeDrawing nearest = CreateHiddenNearestPatternScene(png, 3D, 3D);
            nearest.AddImageWithInterpolation(png, "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 3D, 3D)), interpolate: false);
            OfficeDrawing grouped = new OfficeDrawing(4D, 3D);
            if (useMask) {
                OfficeDrawing source = new OfficeDrawing(3D, 3D);
                OfficeShape rectangle = OfficeShape.Rectangle(3D, 3D);
                rectangle.FillColor = OfficeColor.Red;
                rectangle.StrokeColor = null;
                source.AddShape(rectangle, 0D, 0D);
                grouped.AddEffectDrawing(source, OfficeTransform.Scale(1.1D, 0.9D), blend,
                    new OfficeDrawingSoftMask(nearest, OfficeSoftMaskMode.Luminosity));
            } else {
                grouped.AddEffectDrawing(nearest, OfficeTransform.Scale(1.1D, 0.9D), blend);
            }
            OfficeDrawing direct = new OfficeDrawing(4D, 3D);
            direct.AddImageWithInterpolation(OfficePngWriter.Encode(expectedPixels), "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 3.3D, 2.7D)), interpolate: false);

            Assert.Equal(OfficeDrawingRasterRenderer.Render(direct).GetPixels(),
                OfficeDrawingRasterRenderer.Render(grouped).GetPixels());
        }

        private static OfficeDrawing CreateHiddenNearestPatternScene(byte[] png, double width, double height) {
            OfficeDrawing clipped = new OfficeDrawing(2D, 1D);
            clipped.AddImageWithInterpolation(png, "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)), interpolate: false);
            OfficeDrawing hiddenTile = new OfficeDrawing(2D, 1D);
            hiddenTile.AddClippedDrawing(clipped, 0D, 0D, OfficeClipPath.Rectangle(2D, 1D), 3D, 0D);
            OfficeTransform minified = OfficeTransform.Scale(1D / 1024D, 1D);
            OfficeDrawing nestedTile = new OfficeDrawing(2D, 1D);
            nestedTile.AddTilingPattern(hiddenTile, new OfficeImagePlacement(0D, 0D, 2D, 1D),
                2D, 1D, repeatX: true, repeatY: false, transform: minified);
            OfficeDrawing drawing = new OfficeDrawing(width, height);
            drawing.AddTilingPattern(nestedTile, new OfficeImagePlacement(0D, 0D, 2D, 1D),
                2D, 1D, repeatX: true, repeatY: false, transform: minified);
            return drawing;
        }
    }
}
