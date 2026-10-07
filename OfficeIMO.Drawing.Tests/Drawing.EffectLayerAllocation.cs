using System;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class DrawingEffectLayerAllocationTests {
        [Theory]
        [InlineData(0, OfficeBlendMode.Normal)]
        [InlineData(1, OfficeBlendMode.Multiply)]
        [InlineData(2, OfficeBlendMode.Normal)]
        public void NoninvertibleEffectsDoNotAllocateInvisibleLayers(int kind, OfficeBlendMode blend) {
            OfficeDrawing inner = Solid(1D, 1D);
            OfficeTransform transform = kind == 0 ? OfficeTransform.Scale(0D, 1000D)
                : kind == 1 ? OfficeTransform.Scale(1000D, 0D)
                : new OfficeTransform(1D, 2D, 2D, 4D, 0D, 0D);
            OfficeDrawing drawing = new OfficeDrawing(10D, 10D);
            drawing.AddEffectDrawing(inner, transform, blend, new OfficeDrawingSoftMask(inner));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 100L });

            Assert.All(image.GetPixels(), value => Assert.Equal((byte)0, value));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void NonuniformSvgViewportAllocatesAndSamplesEachAxis(bool vertical) {
            string svg = vertical
                ? "<svg xmlns='http://www.w3.org/2000/svg' width='1' height='1000' viewBox='0 0 1000 1' preserveAspectRatio='none'><rect width='1000' height='1' fill='white'/><rect width='1000' height='.25'/></svg>"
                : "<svg xmlns='http://www.w3.org/2000/svg' width='1000' height='1' viewBox='0 0 1 1000' preserveAspectRatio='none'><rect width='1' height='1000' fill='white'/><rect width='.25' height='1000'/></svg>";
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int loss));
            Assert.Equal(0, loss);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing!,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 10_000L });

            Assert.Equal(vertical ? 1 : 1000, image.Width);
            Assert.Equal(vertical ? 1000 : 1, image.Height);
            for (int index = 0; index < 1000; index++)
                Assert.Equal(index < 250 ? OfficeColor.Black : OfficeColor.White,
                    image.GetPixel(vertical ? 0 : index, vertical ? index : 0));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void RotatedOrReflectedAnisotropicEffectsRetainDestinationFootprint(bool rotate) {
            OfficeTransform transform = OfficeTransform.Scale(1000D, .001D).Then(rotate
                ? OfficeTransform.RotateDegrees(90D).Then(OfficeTransform.Translate(1D, 0D))
                : OfficeTransform.Scale(-1D, 1D).Then(OfficeTransform.Translate(1000D, 0D)));
            OfficeDrawing drawing = new OfficeDrawing(rotate ? 1D : 1000D, rotate ? 1000D : 1D);
            drawing.AddEffectDrawing(Solid(1D, 1000D), transform, OfficeBlendMode.Multiply);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 10_000L });

            for (int index = 0; index < 1000; index++)
                Assert.Equal(OfficeColor.Black, image.GetPixel(rotate ? 0 : index, rotate ? index : 0));
        }

        [Fact]
        public void NestedMaskedAnisotropicEffectSharesActualAllocationBudget() {
            OfficeDrawing mask = Solid(.5D, 1000D);
            OfficeDrawing masked = new OfficeDrawing(1000D, 1D);
            masked.AddEffectDrawing(Solid(1D, 1000D), OfficeTransform.Scale(1000D, .001D), OfficeBlendMode.Normal,
                new OfficeDrawingSoftMask(mask, OfficeSoftMaskMode.Alpha));
            OfficeDrawing drawing = new OfficeDrawing(1000D, 1D);
            drawing.AddEffectDrawing(masked, OfficeTransform.Identity);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 10_000L });

            Assert.Equal(OfficeColor.Black, image.GetPixel(100, 0));
            Assert.Equal(OfficeColor.Transparent, image.GetPixel(900, 0));
            Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 3_000L }));
        }

        [Fact]
        public void AnisotropicClippedGradientAndTextRetainTheirOwnersAndCancellation() {
            RecordingShaper shaper = new RecordingShaper();
            OfficeDrawing child = new OfficeDrawing(1D, 1000D) { TextShapingProvider = shaper, TextShapingLanguage = "ar-SA" };
            OfficeShape rectangle = OfficeShape.Rectangle(1D, 1000D);
            rectangle.FillColor = OfficeColor.Red;
            rectangle.FillGradient = new OfficeLinearGradient(0D, 0D, 1D, 0D,
                new OfficeGradientStop(0D, OfficeColor.Red), new OfficeGradientStop(1D, OfficeColor.Blue));
            rectangle.StrokeColor = null;
            child.AddShape(rectangle, 0D, 0D);
            child.AddPositionedText("A", 0D, 0D, 1D, 1000D, new OfficeFontInfo("Arial", 1D), textAdvanceWidth: 1D);
            OfficeDrawing clipped = new OfficeDrawing(1D, 1000D);
            clipped.AddClippedDrawing(child, 0D, 0D, OfficeClipPath.Rectangle(.5D, 1000D));
            OfficeDrawing drawing = new OfficeDrawing(1000D, 1D);
            drawing.AddEffectDrawing(clipped, OfficeTransform.Scale(1000D, .001D));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 10_000L });

            Assert.True(image.GetPixel(100, 0).R > image.GetPixel(100, 0).B);
            Assert.Equal(OfficeColor.Transparent, image.GetPixel(900, 0));
            Assert.True(shaper.Calls > 0);
            using CancellationTokenSource cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            Assert.Throws<OperationCanceledException>(() => OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { CancellationToken = cancelled.Token }));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void AnisotropicImageAndPatternUseActualIntermediateAxes(bool tile) {
            OfficeDrawing inner = new OfficeDrawing(1D, 1000D);
            if (tile) {
                inner.AddTilingPattern(Solid(1D, 1000D), new OfficeImagePlacement(0D, 0D, 1D, 1000D), 1D, 1000D);
            } else {
                byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Black));
                inner.AddImage(png, "image/png", new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 1D, 1000D)));
            }
            OfficeDrawing drawing = new OfficeDrawing(1000D, 1D);
            drawing.AddEffectDrawing(inner, OfficeTransform.Scale(1000D, .001D));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 10_000L });

            for (int index = 0; index < 1000; index++) Assert.Equal(OfficeColor.Black, image.GetPixel(index, 0));
        }

        [Fact]
        public void AnisotropicTransformedTextAllocatesItsPhysicalFrameAndKeepsShaping() {
            RecordingShaper shaper = new RecordingShaper();
            OfficeDrawing inner = new OfficeDrawing(1D, 1000D) { TextShapingProvider = shaper, TextShapingLanguage = "ar-SA" };
            inner.AddPositionedText("A", 0D, 0D, 1D, 1000D,
                new OfficeImageFrameTransform(0D, .5D, 500D, flipHorizontal: true),
                new OfficeFontInfo("Arial", 1000D), textAdvanceWidth: 1D);
            OfficeDrawing drawing = new OfficeDrawing(1000D, 1D);
            drawing.AddEffectDrawing(inner, OfficeTransform.Scale(1000D, .001D));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 10_000L });

            Assert.True(shaper.Calls > 0);
            Assert.Contains(image.GetPixels(), value => value != 0);
        }

        private static OfficeDrawing Solid(double width, double height) {
            OfficeDrawing drawing = new OfficeDrawing(width, height);
            OfficeShape rectangle = OfficeShape.Rectangle(width, height);
            rectangle.FillColor = OfficeColor.Black;
            rectangle.StrokeColor = null;
            drawing.AddShape(rectangle, 0D, 0D);
            return drawing;
        }

        private sealed class RecordingShaper : IOfficeTextShapingProvider {
            internal int Calls { get; private set; }
            public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) { Calls++; return null; }
        }
    }
}
