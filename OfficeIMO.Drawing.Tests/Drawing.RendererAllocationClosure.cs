using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class DrawingRendererAllocationClosureTests {
        [Theory]
        [InlineData("image", false)]
        [InlineData("image", true)]
        [InlineData("image-pattern", false)]
        [InlineData("image-pattern", true)]
        [InlineData("vector-pattern", false)]
        [InlineData("vector-pattern", true)]
        [InlineData("positioned-text", false)]
        [InlineData("positioned-text", true)]
        [InlineData("vertical-text", false)]
        [InlineData("vertical-text", true)]
        public void ClippedRendererContentDoesNotConsumeInvisibleSurfaceBudget(string kind, bool polygon) {
            OfficeDrawing content = Content(kind);
            OfficeClipPath clip = polygon ? OfficeClipPath.Path(OfficePathCommand.MoveTo(0D, 0D),
                OfficePathCommand.LineTo(1D, 0D), OfficePathCommand.LineTo(1D, 1000D),
                OfficePathCommand.LineTo(0D, 1000D), OfficePathCommand.Close()) : OfficeClipPath.Rectangle(1D, 1000D);
            OfficeDrawing inner = new OfficeDrawing(10D, 1000D).AddClippedDrawing(content, 0D, 0D, clip, 5D, 0D);
            OfficeDrawing drawing = new OfficeDrawing(1000D, 1D).AddEffectDrawing(inner, OfficeTransform.Scale(100D, .001D));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 1000L });

            Assert.All(image.GetPixels(), value => Assert.Equal((byte)0, value));
        }

        [Theory]
        [InlineData("image")]
        [InlineData("image-pattern")]
        [InlineData("vector-pattern")]
        [InlineData("positioned-text")]
        [InlineData("vertical-text")]
        public void VisibleRendererContentStillChargesSharedIntermediateBudget(string kind) {
            OfficeDrawing inner = Content(kind);
            OfficeDrawing drawing = new OfficeDrawing(1000D, 1D).AddEffectDrawing(inner, OfficeTransform.Scale(1000D, .001D));

            Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 1000L }));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void RotatedAnisotropicVectorPatternRetainsNarrowDestinationDetail(bool reflected) {
            var tile = new OfficeDrawing(100D, 1D);
            AddRectangle(tile, 0D, 0D, 100D, 1D, OfficeColor.White);
            AddRectangle(tile, 0D, 0D, 100D, .25D, OfficeColor.Black);
            OfficeTransform transform = reflected ? new OfficeTransform(0D, 1D, 1D, 0D, 0D, 0D)
                : new OfficeTransform(0D, 1D, -1D, 0D, 1D, 0D);
            var inner = new OfficeDrawing(1D, 100D);
            inner.AddTilingPattern(tile, new OfficeImagePlacement(0D, 0D, 1D, 100D), 100D, 1D,
                repeatX: false, repeatY: false, transform: transform);
            OfficeDrawing drawing = new OfficeDrawing(100D, 1D).AddEffectDrawing(inner, OfficeTransform.Scale(100D, .01D));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 200L });

            for (int x = 0; x < 100; x++) Assert.Equal((reflected ? x < 25 : x >= 75) ? OfficeColor.Black : OfficeColor.White,
                image.GetPixel(x, 0));
            Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 199L }));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PatternGapsSkipImageDecodeAndVectorTileAllocation(bool imagePattern) {
            var drawing = new OfficeDrawing(10D, 10D);
            var area = new OfficeImagePlacement(0D, 0D, 10D, 10D);
            if (imagePattern) drawing.AddImagePattern(Svg(1000D, 1000D), "image/svg+xml",
                new OfficeImagePatternLayout(area, new OfficeImagePlacement(20D, 0D, 1000D, 1000D), false, false));
            else drawing.AddTilingPattern(new OfficeDrawing(1000D, 1000D), area, 1000D, 1000D,
                repeatX: false, repeatY: false, originX: 20D);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 100L });

            Assert.All(image.GetPixels(), value => Assert.Equal((byte)0, value));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void FractionalPatternAreaKeepsItsCoveringPixelClip(bool imagePattern) {
            var drawing = new OfficeDrawing(1D, 1D);
            var area = new OfficeImagePlacement(0D, 0D, .1D, .1D);
            var placement = new OfficeImagePlacement(0D, 0D, 1D, 1D);
            if (imagePattern) drawing.AddImagePattern(OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Black)),
                "image/png", new OfficeImagePatternLayout(area, placement, false, false));
            else {
                var tile = new OfficeDrawing(1D, 1D);
                AddRectangle(tile, 0D, 0D, 1D, 1D, OfficeColor.Black);
                drawing.AddTilingPattern(tile, area, 1D, 1D, repeatX: false, repeatY: false);
            }

            Assert.Equal(OfficeColor.Black, OfficeDrawingRasterRenderer.Render(drawing).GetPixel(0, 0));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void EncodedReflectedImagesKeepIncludedFractionalPolygonEdges(bool vertical) {
            byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Black));
            var projection = new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 1D, 1D),
                rotationCenterX: vertical ? .5D : .25D, rotationCenterY: vertical ? .25D : .5D,
                flipHorizontal: !vertical, flipVertical: vertical);
            double left = vertical ? 0D : .5D, top = vertical ? .5D : 0D;
            OfficeClipPath clip = OfficeClipPath.Path(OfficePathCommand.MoveTo(0D, 0D),
                OfficePathCommand.LineTo(4D, 0D), OfficePathCommand.LineTo(4D, 4D),
                OfficePathCommand.LineTo(0D, 4D), OfficePathCommand.Close());
            var drawing = new OfficeDrawing(8D, 8D).AddClippedImage(png, "image/png", projection, left, top, clip);

            Assert.Equal(OfficeColor.Black, OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 64L }).GetPixel(0, 0));
        }

        [Theory]
        [InlineData("image")]
        [InlineData("image-pattern")]
        [InlineData("vector-pattern")]
        [InlineData("positioned-text")]
        [InlineData("vertical-text")]
        public void EmptyRendererClipCannotResurrectAnIntermediateAllocation(string kind) {
            var inner = new OfficeDrawing(10D, 1000D).AddClippedDrawing(Content(kind), 5D, 0D,
                OfficeClipPath.Rectangle(.002D, 1000D));
            var drawing = new OfficeDrawing(1000D, 1D).AddEffectDrawing(inner, OfficeTransform.Scale(100D, .001D));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 1000L });

            Assert.All(image.GetPixels(), value => Assert.Equal((byte)0, value));
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void HiddenEncodedContentSkipsRequiredCallerDecode(bool pattern, bool zeroOpacity) {
            var drawing = new OfficeDrawing(8D, 8D);
            byte[] bytes = { 1, 2, 3 };
            var codec = new RecordingCodec();
            var child = new OfficeDrawing(8D, 8D);
            if (pattern) child.AddImagePattern(bytes, "application/test-raster",
                new OfficeImagePatternLayout(new OfficeImagePlacement(0D, 0D, 8D, 8D),
                    new OfficeImagePlacement(0D, 0D, 8D, 8D)), opacity: zeroOpacity ? 0D : .5D);
            else child.AddImage(bytes, "application/test-raster",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 8D, 8D)), opacity: zeroOpacity ? 0D : .5D);
            if (zeroOpacity) drawing = child;
            else drawing.AddClippedDrawing(child, 0D, 0D, OfficeClipPath.Rectangle(8D, 8D), 20D, 0D);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { ThrowOnImageDecodeFailure = true, ImageCodec = codec });

            Assert.Equal(0, codec.Calls);
            Assert.All(image.GetPixels(), value => Assert.Equal((byte)0, value));
        }

        [Fact]
        public void PositionedTextCullIncludesMeasuredInkBeyondItsFrame() {
            var drawing = new OfficeDrawing(8D, 8D).AddFont("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
            drawing.AddPositionedText("A", 0D, 0D, 1D, 8D,
                new OfficeImageFrameTransform(0D, 5D, 4D, flipHorizontal: true), new OfficeFontInfo("Ink", 8D), textAdvanceWidth: 8D);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 256L });

            Assert.Contains(image.GetPixels(), value => value != 0);
        }

        [Fact]
        public void VisibleTransformedVerticalTextKeepsItsLayerAndSharedCharge() {
            var child = new OfficeDrawing(20D, 20D).AddFont("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
            child.AddVerticalText("A", 0D, 0D, 20D, 20D, new OfficeFontInfo("Ink", 8D));
            var drawing = new OfficeDrawing(20D, 20D).AddDrawing(child, 0D, 0D, new OfficeImageFrameTransform(90D, 10D, 10D));
            drawing.TextShapingProvider = new VerticalProvider();
            var budget = new OfficeRasterTransformedTextBudget();

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 400L, TransformedTextBudget = budget });

            Assert.Contains(image.GetPixels(), value => value != 0);
            Assert.Equal(400L, budget.IntermediatePixels);
        }

        [Fact]
        public void WrappedRotatedTextPaintsDirectlyWithoutAnIntermediateLayer() {
            var child = new OfficeDrawing(10D, 1000D).AddFont("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
            child.AddText("A", 0D, 0D, 1D, 1000D, new OfficeFontInfo("Ink", 1D),
                rotationDegrees: 90D, rotationCenterX: .5D, rotationCenterY: 500D, wrapText: true);
            var clipped = new OfficeDrawing(10D, 1000D).AddClippedDrawing(child, 0D, 0D, OfficeClipPath.Rectangle(1D, 1000D), 5D, 0D);
            var drawing = new OfficeDrawing(1000D, 1D).AddEffectDrawing(clipped, OfficeTransform.Scale(100D, .001D));
            var budget = new OfficeRasterTransformedTextBudget();

            OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 1000L, TransformedTextBudget = budget });

            Assert.Equal(1000L, budget.IntermediatePixels);
        }

        private sealed class VerticalProvider : IOfficeTextShapingProvider {
            public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) => new OfficeTextShapingResult(new[] {
                new OfficeShapedGlyph(1, "A", 0, advanceWidth: 0, advanceHeight: -1000, offsetX: 0, offsetY: 0)
            }, OfficeTextDirection.TopToBottom);
        }

        private sealed class RecordingCodec : IOfficeRasterImageCodec {
            internal int Calls { get; private set; }
            public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
                Calls++;
                image = new OfficeRasterImage(1, 1, OfficeColor.Black);
                return true;
            }
        }

        private static OfficeDrawing Content(string kind) {
            var child = new OfficeDrawing(1D, 1000D).AddFont("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
            var placement = new OfficeImagePlacement(0D, 0D, 1D, 1000D);
            if (kind == "image") child.AddImage(Svg(1D, 1000D), "image/svg+xml", new OfficeImageProjection(placement));
            else if (kind == "image-pattern") child.AddImagePattern(Svg(1D, 1000D), "image/svg+xml", new OfficeImagePatternLayout(placement, placement), opacity: .5D);
            else if (kind == "vector-pattern") {
                var tile = new OfficeDrawing(1D, 1000D);
                AddRectangle(tile, 0D, 0D, 1D, 1000D, OfficeColor.Black);
                child.AddTilingPattern(tile, placement, 1D, 1000D);
            }
            else if (kind == "positioned-text") child.AddPositionedText("A", 0D, 0D, 1D, 1000D,
                new OfficeImageFrameTransform(0D, .5D, 500D, flipHorizontal: true), new OfficeFontInfo("Ink", 1D), textAdvanceWidth: 1D);
            else {
                var vertical = new OfficeDrawing(1D, 1000D).AddFont("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
                vertical.AddVerticalText("A", 0D, 0D, 1D, 1000D, new OfficeFontInfo("Ink", 1D));
                child.AddDrawing(vertical, 0D, 0D, new OfficeImageFrameTransform(0D, .5D, 500D, flipHorizontal: true));
            }
            return child;
        }

        private static byte[] Svg(double width, double height) => Encoding.UTF8.GetBytes(
            $"<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}'><rect width='{width}' height='{height}' fill='black'/></svg>");

        private static void AddRectangle(OfficeDrawing drawing, double x, double y, double width, double height, OfficeColor color) {
            OfficeShape shape = OfficeShape.Rectangle(width, height);
            shape.FillColor = color;
            shape.StrokeColor = null;
            drawing.AddShape(shape, x, y);
        }
    }
}
