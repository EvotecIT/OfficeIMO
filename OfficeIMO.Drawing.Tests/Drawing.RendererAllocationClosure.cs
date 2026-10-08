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

        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void UnsupportedVerticalFrameRetainsVisibleFallbackInk(bool vertical, bool horizontalProvider) {
            OfficeDrawing drawing = VerticalFrame(vertical);
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { TextShapingProvider = horizontalProvider ? new HorizontalProvider() : null });

            Assert.Equal(vertical ? (byte)255 : (byte)179, image.GetPixel(vertical ? 3 : 7, 3).A);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void SupportedHiddenVerticalFrameSkipsLayerAndFallback(bool vertical) {
            var provider = new VerticalProvider();
            var budget = new OfficeRasterTransformedTextBudget();
            budget.ChargeIntermediateSurfacePixels(64L, 64L);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(VerticalFrame(vertical),
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 64L, TransformedTextBudget = budget, TextShapingProvider = provider });

            Assert.True(provider.Calls > 0);
            Assert.All(image.GetPixels(), value => Assert.Equal((byte)0, value));
            Assert.Equal(64L, budget.IntermediatePixels);
        }

        [Fact]
        public void HiddenVerticalCapabilityCheckObservesProviderCancellation() {
            using var cancellation = new System.Threading.CancellationTokenSource();
            var provider = new VerticalProvider(cancellation.Cancel);

            Assert.ThrowsAny<OperationCanceledException>(() => OfficeDrawingRasterRenderer.Render(VerticalFrame(false),
                new OfficeDrawingRasterRenderOptions { TextShapingProvider = provider, CancellationToken = cancellation.Token }));
        }

        [Theory]
        [InlineData(false, "effect")]
        [InlineData(true, "effect")]
        [InlineData(false, "pattern")]
        [InlineData(true, "pattern")]
        [InlineData(false, "nested-effect")]
        [InlineData(true, "nested-effect")]
        [InlineData(false, "mask")]
        [InlineData(true, "mask")]
        public void ReflectedNearestBoundarySurvivesSharedSamplingInspection(bool vertical, string route) {
            var scene = new OfficeDrawing(1D, 1D);
            scene.AddClippedImageWithInterpolation(OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Black)), "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 1D, 1D),
                    rotationCenterX: vertical ? .5D : .25D, rotationCenterY: vertical ? .25D : .5D,
                    flipHorizontal: !vertical, flipVertical: vertical), false, vertical ? 0D : .5D, vertical ? .5D : 0D,
                OfficeClipPath.Rectangle(vertical ? 1D : .5D, vertical ? .5D : 1D));
            Assert.Equal(OfficeColor.Black, OfficeDrawingRasterRenderer.Render(scene).GetPixel(0, 0));
            OfficeTransform transform = OfficeTransform.Scale(vertical ? 1D : 2D, vertical ? 2D : 1D);
            var drawing = new OfficeDrawing(vertical ? 1D : 2D, vertical ? 2D : 1D);
            if (route == "pattern") drawing.AddTilingPattern(scene, new OfficeImagePlacement(0D, 0D, drawing.Width, drawing.Height),
                1D, 1D, repeatX: false, repeatY: false, transform: transform);
            else if (route == "nested-effect") drawing.AddEffectDrawing(new OfficeDrawing(1D, 1D).AddEffectDrawing(scene, OfficeTransform.Identity), transform);
            else if (route == "mask") {
                var solid = new OfficeDrawing(1D, 1D);
                AddRectangle(solid, 0D, 0D, 1D, 1D, OfficeColor.Black);
                drawing.AddEffectDrawing(solid, transform, OfficeBlendMode.Normal, new OfficeDrawingSoftMask(scene));
            } else drawing.AddEffectDrawing(scene, transform);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            for (int index = 0; index < 2; index++) Assert.Equal(OfficeColor.Black, image.GetPixel(vertical ? 0 : index, vertical ? index : 0));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void DisjointAndEmptyNearestClipsKeepVisibleInterpolatedSampling(bool empty) {
            var source = new OfficeRasterImage(2, 1, OfficeColor.Black);
            source.SetPixel(1, 0, OfficeColor.White);
            byte[] png = OfficePngWriter.Encode(source);
            var projection = new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D));
            var tile = new OfficeDrawing(2D, 1D).AddImageWithInterpolation(png, "image/png", projection, true);
            var hidden = new OfficeDrawing(2D, 1D).AddImageWithInterpolation(png, "image/png", projection, false);
            tile.AddClippedDrawing(hidden, 0D, 0D, empty ? OfficeClipPath.Empty() : OfficeClipPath.Rectangle(1D, 1D),
                empty ? 0D : 10D, 0D);
            var drawing = new OfficeDrawing(4D, 1D).AddTilingPattern(tile, new OfficeImagePlacement(0D, 0D, 4D, 1D),
                2D, 1D, repeatX: false, repeatY: false, transform: OfficeTransform.Scale(2D, 1D));

            OfficeColor boundary = OfficeDrawingRasterRenderer.Render(drawing).GetPixel(1, 0);

            Assert.InRange(boundary.R, (byte)1, (byte)254);
            Assert.Equal(boundary.R, boundary.G);
            Assert.Equal(boundary.R, boundary.B);
        }

        [Theory]
        [InlineData(.018D)]
        [InlineData(.13D)]
        public void SubpixelPeriodicTilesRetainOpaqueCoverage(double density) {
            var tile = new OfficeDrawing(10D, 10D);
            AddRectangle(tile, 0D, 0D, 10D, 10D, OfficeColor.Black);
            var drawing = new OfficeDrawing(10D, 10D).AddTilingPattern(tile,
                new OfficeImagePlacement(0D, 0D, 10D, 10D), 10D, 10D,
                transform: OfficeTransform.Scale(density, density));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

            for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++)
                Assert.Equal(OfficeColor.Black, image.GetPixel(x, y));
        }

        [Fact]
        public void FractionalTileDimensionsRetainTheirPlannedPixelBudget() {
            var tile = new OfficeDrawing(.3D, 1D);
            AddRectangle(tile, 0D, 0D, .3D, 1D, OfficeColor.Black);
            var drawing = new OfficeDrawing(7D, 1D).AddTilingPattern(tile,
                new OfficeImagePlacement(0D, 0D, 7D, 1D), .3D, 1D,
                transform: OfficeTransform.Scale(22D, 1D));

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing,
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 7L });

            Assert.Equal(7, image.Width);
            Assert.Equal(1, image.Height);
            for (int x = 0; x < image.Width; x++) Assert.Equal(OfficeColor.Black, image.GetPixel(x, 0));
        }

        [Theory]
        [InlineData(5.8D, false)]
        [InlineData(6.2D, true)]
        public void FractionalInterpolatedTileUsesItsPaintedBoundsBeforeAllocation(double offsetX, bool visible) {
            var tile = new OfficeDrawing(1D, 1000D);
            AddRectangle(tile, 0D, 0D, 1D, 1000D, OfficeColor.Black);
            var child = new OfficeDrawing(10D, 1D).AddTilingPattern(tile,
                new OfficeImagePlacement(0D, 0D, 10D, 1D), 1D, 1000D,
                repeatX: false, repeatY: false, transform: new OfficeTransform(.6D, 0D, 0D, 1D, offsetX, 0D));
            var drawing = new OfficeDrawing(10D, 1D).AddClippedDrawing(child, 6D, 0D,
                OfficeClipPath.Rectangle(1D, 1D), -6D, 0D);
            var options = new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 10L };

            if (visible) {
                Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing, options));
            } else {
                OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, options);
                for (int x = 0; x < image.Width; x++) Assert.Equal(0, image.GetPixel(x, 0).A);
            }
        }

        [Theory]
        [InlineData(95L, false)]
        [InlineData(96L, true)]
        [InlineData(123L, true)]
        [InlineData(124L, true)]
        public void OptionalTextPaddingUsesRemainingIntermediateBudget(long maximumPixels, bool succeeds) {
            var drawing = PositionedFrame();
            drawing = new OfficeDrawing(8D, 8D).AddImage(OfficePngWriter.Encode(new OfficeRasterImage(8, 8, OfficeColor.White)),
                "image/png", new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 8D, 8D))).AddDrawing(drawing, 0D, 0D);
            var options = new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = maximumPixels };

            if (succeeds) Assert.NotEqual(OfficeColor.White, OfficeDrawingRasterRenderer.Render(drawing, options).GetPixel(7, 0));
            else Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing, options));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void OptionalTextPaddingHonorsRememberedAndCumulativeTextCeilings(bool cumulativeText) {
            var budget = new OfficeRasterTransformedTextBudget();
            if (cumulativeText) {
                var canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1));
                canvas.ShareTransformedTextBudget(budget);
                canvas.ChargeTransformedTextIntermediatePixels(63_999_968L, 128_000_000L);
            } else budget.ChargeIntermediateSurfacePixels(64L, 96L);

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(PositionedFrame(),
                new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = cumulativeText ? 128_000_000L : 200L, TransformedTextBudget = budget });

            Assert.True(image.GetPixel(7, 0).A > 0);
            Assert.Equal(cumulativeText ? 64_000_000L : 96L, budget.IntermediatePixels);
        }

        [Theory]
        [InlineData(111L, false)]
        [InlineData(112L, true)]
        public void OptionalTextPaddingUsesRoundedAnisotropicLayerPixels(long maximumPixels, bool succeeds) {
            var drawing = new OfficeDrawing(8D, 8D).AddImage(OfficePngWriter.Encode(new OfficeRasterImage(8, 8, OfficeColor.White)),
                "image/png", new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 8D, 8D)));
            drawing.AddEffectDrawing(PositionedFrame(), OfficeTransform.Scale(.7D, 1D));
            var options = new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = maximumPixels };

            if (succeeds) Assert.NotEqual(OfficeColor.White, OfficeDrawingRasterRenderer.Render(drawing, options).GetPixel(5, 0));
            else Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing, options));
        }

        private static OfficeDrawing VerticalFrame(bool vertical) {
            var child = new OfficeDrawing(8D, 8D).AddFont("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
            child.AddVerticalText("A", 0D, 0D, vertical ? 8D : 1D, vertical ? 1D : 8D, new OfficeFontInfo("Ink", 8D));
            return new OfficeDrawing(8D, 8D).AddDrawing(child, 0D, 0D,
                new OfficeImageFrameTransform(0D, vertical ? 4D : 4.5D, vertical ? 4.5D : 4D,
                    flipHorizontal: !vertical, flipVertical: vertical));
        }

        private static OfficeDrawing PositionedFrame() => new OfficeDrawing(8D, 4D).AddFont("Ink", ManagedTextShapingTestAssets.CreateFont('A'))
            .AddPositionedText("A", 0D, 0D, 8D, 4D, new OfficeImageFrameTransform(0D, 4D, 2D, flipHorizontal: true),
                new OfficeFontInfo("Ink", 1D), textAdvanceWidth: 1D);

        private sealed class VerticalProvider : IOfficeTextShapingProvider {
            private readonly Action? _onShape;
            internal int Calls { get; private set; }
            internal VerticalProvider(Action? onShape = null) => _onShape = onShape;
            public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
                Calls++;
                _onShape?.Invoke();
                return request.Direction == OfficeTextDirection.TopToBottom ? new OfficeTextShapingResult(new[] {
                    new OfficeShapedGlyph(1, "A", 0, advanceWidth: 0, advanceHeight: -1000, offsetX: 0, offsetY: 0)
                }, OfficeTextDirection.TopToBottom) : new HorizontalProvider().ShapeText(request);
            }
        }

        private sealed class HorizontalProvider : IOfficeTextShapingProvider {
            public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) => new OfficeTextShapingResult(new[] {
                new OfficeShapedGlyph(1, "A", 0, advanceWidth: 500, advanceHeight: 0, offsetX: 0, offsetY: 0)
            }, OfficeTextDirection.LeftToRight);
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
