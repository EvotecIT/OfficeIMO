using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class DrawingRasterFeedbackTests {
        [Theory]
        [InlineData(false, false, "line")]
        [InlineData(false, true, "line")]
        [InlineData(true, false, "line")]
        [InlineData(true, true, "line")]
        [InlineData(false, false, "path")]
        [InlineData(false, true, "path")]
        [InlineData(true, false, "path")]
        [InlineData(true, true, "path")]
        [InlineData(false, false, "rect")]
        [InlineData(false, true, "rect")]
        [InlineData(true, false, "rect")]
        [InlineData(true, true, "rect")]
        public void CompressedSvgViewportRetainsStrokeOnlyGeometry(bool horizontal, bool transformed, string kind) {
            string transform = transformed ? " transform='translate(0 100)'" : "";
            int position = transformed ? 700 : 800;
            string geometry = kind == "line" ? $"<line x1='20' y1='{position}' x2='80' y2='{position}'{transform}/>"
                : kind == "path" ? $"<path d='M20 {position} L80 {position}'{transform}/>"
                : $"<rect x='20' y='{position - 100}' width='60' height='100'{transform}/>";
            if (horizontal) geometry = "<g transform='matrix(0 1 1 0 0 0)'>" + geometry + "</g>";
            string viewBox = horizontal ? "0 0 1000 100" : "0 0 100 1000";
            string svg = $"<svg xmlns='http://www.w3.org/2000/svg' width='100' height='100' viewBox='{viewBox}' preserveAspectRatio='none'>"
                + "<g fill='none' stroke='black' stroke-width='20'>" + geometry + "</g></svg>";
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int loss));
            Assert.Equal(0, loss);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing!);

            Assert.Equal(OfficeColor.Black, image.GetPixel(horizontal ? 80 : 50, horizontal ? 50 : 80));
            Assert.Equal(OfficeColor.Transparent, image.GetPixel(horizontal ? 84 : 50, horizontal ? 50 : 84));
        }

        [Theory]
        [InlineData(0, OfficeBlendMode.Normal)]
        [InlineData(1, OfficeBlendMode.Multiply)]
        [InlineData(2, OfficeBlendMode.Normal)]
        [InlineData(3, OfficeBlendMode.Multiply)]
        public void OffCanvasMagnifiedEffectsDoNotAllocateInvisibleLayers(int direction, OfficeBlendMode blend) {
            OfficeTransform transform = OfficeTransform.Scale(1000D, 1000D).Then(direction switch {
                0 => OfficeTransform.Translate(20D, 0D),
                1 => OfficeTransform.Translate(-1020D, 0D),
                2 => OfficeTransform.RotateDegrees(90D).Then(OfficeTransform.Translate(0D, 20D)),
                _ => OfficeTransform.Scale(-1D, -1D).Then(OfficeTransform.Translate(0D, -20D))
            });
            OfficeDrawing source = Solid(1D, 1D);
            OfficeDrawing drawing = new OfficeDrawing(10D, 10D)
                .AddEffectDrawing(source, transform, blend, new OfficeDrawingSoftMask(source));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 100L });

            Assert.All(image.GetPixels(), value => Assert.Equal((byte)0, value));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ClippedMagnifiedEffectsDoNotAllocateInvisibleLayers(bool pathClip) {
            OfficeDrawing content = new OfficeDrawing(10D, 10D)
                .AddEffectDrawing(Solid(1D, 1D), OfficeTransform.Scale(1000D, 1000D).Then(OfficeTransform.Translate(6D, 0D)));
            OfficeClipPath clip = pathClip ? OfficeClipPath.Path(OfficePathCommand.MoveTo(0D, 0D),
                OfficePathCommand.LineTo(5D, 0D), OfficePathCommand.LineTo(5D, 10D),
                OfficePathCommand.LineTo(0D, 10D), OfficePathCommand.Close()) : OfficeClipPath.Rectangle(5D, 10D);
            OfficeDrawing drawing = new OfficeDrawing(10D, 10D).AddClippedDrawing(content, 0D, 0D, clip);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 100L });

            Assert.All(image.GetPixels(), value => Assert.Equal((byte)0, value));
        }

        [Fact]
        public void NestedCompressedCanvasCullsEffectsInPhysicalDestinationCoordinates() {
            OfficeDrawing inner = new OfficeDrawing(100D, 1000D)
                .AddEffectDrawing(Solid(1D, 1D), OfficeTransform.Scale(1000D, 1000D).Then(OfficeTransform.Translate(0D, 1001D)));
            OfficeDrawing drawing = new OfficeDrawing(100D, 100D)
                .AddEffectDrawing(inner, OfficeTransform.Scale(1D, .1D));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 20_000L });

            Assert.All(image.GetPixels(), value => Assert.Equal((byte)0, value));
        }

        [Fact]
        public void RoundedEffectSurfaceRetainsVisibleEdgeBeyondLogicalSize() {
            OfficeShape rectangle = OfficeShape.Rectangle(1.1D, 1D);
            rectangle.FillColor = null;
            rectangle.StrokeColor = OfficeColor.Black;
            rectangle.StrokeWidth = .8D;
            OfficeDrawing inner = new OfficeDrawing(1.1D, 1D).AddShape(rectangle, 0D, 0D);
            OfficeDrawing drawing = new OfficeDrawing(1D, 1D)
                .AddEffectDrawing(inner, OfficeTransform.Scale(2D, 2D).Then(OfficeTransform.Translate(-2.3D, 0D)));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 6L });

            Assert.True(image.GetPixel(0, 0).A > 0);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void SvgEffectFringesRemainVisibleWhenTheirSourceGeometryIsOffCanvas(bool shadow) {
            string primitive = shadow ? "<feDropShadow dx='2' dy='0' stdDeviation='1'/>"
                : "<feGaussianBlur stdDeviation='1'/>";
            string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 12 10'><defs>"
                + "<filter id='edge' filterUnits='userSpaceOnUse' x='0' y='0' width='12' height='10'>" + primitive + "</filter></defs>"
                + "<rect x='2' y='2' width='4' height='4' fill='black' filter='url(#edge)'/></svg>";
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int loss));
            Assert.Equal(0, loss);

            OfficeDrawing translated = new OfficeDrawing(10D, 10D).AddEffectDrawing(drawing!, OfficeTransform.Translate(-6D, 0D));
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(translated);

            Assert.True(image.GetPixel(0, 4).A > 0);
        }

        [Theory]
        [InlineData(0, false)]
        [InlineData(1, false)]
        [InlineData(2, false)]
        [InlineData(0, true)]
        [InlineData(1, true)]
        [InlineData(2, true)]
        public void InvisibleImageMinificationDoesNotChargePrefilterSurfaces(int route, bool clipped) {
            OfficeRasterImage source = new OfficeRasterImage(16, 16, OfficeColor.Black);
            OfficeRasterImage destination = new OfficeRasterImage(8, 8);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(destination);
            canvas.ChargeIntermediateSurfacePixels(256L, 256L);
            using IDisposable? clip = clipped ? canvas.PushClipRectangle(0D, 0D, 4D, 8D) : null;
            double x = clipped ? 6D : 20D;

            if (route == 0) canvas.DrawImage(source, new OfficeImageProjection(new OfficeImagePlacement(x, 0D, 8D, 8D)));
            else canvas.DrawAffineImage(source, OfficeTransform.Scale(.5D, .5D).Then(OfficeTransform.Translate(x, 0D)),
                1D, route == 1 ? OfficeBlendMode.Normal : OfficeBlendMode.Multiply);

            Assert.Equal(256L, canvas.TransformedTextBudget.IntermediatePixels);
            Assert.All(destination.GetPixels(), value => Assert.Equal((byte)0, value));
        }

        private static OfficeDrawing Solid(double width, double height) {
            OfficeShape rectangle = OfficeShape.Rectangle(width, height);
            rectangle.FillColor = OfficeColor.Black;
            rectangle.StrokeColor = null;
            return new OfficeDrawing(width, height).AddShape(rectangle, 0D, 0D);
        }
    }
}
