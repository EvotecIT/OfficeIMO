using OfficeIMO.Drawing;
using System;
using System.IO;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class RasterApiFoundationContracts {
        [Fact]
        public void PublicAllocationRejectsOversizedPixelsBeforeAllocating() {
            Assert.Throws<ArgumentException>(() => new OfficeRasterImage(50_000_001, 1));
        }

        [Fact]
        public void SnapshotAndSelfDrawRespectCombinedStorageBeforeMutation() {
            var image = new OfficeRasterImage(10_000, 3_400);
            image.SetPixel(0, 0, OfficeColor.Blue);
            Assert.Throws<ArgumentException>(() => image.Clone());
            Assert.Throws<ArgumentException>(() => image.GetPixels());
            Assert.Throws<ArgumentException>(() => new OfficeRasterCanvas(image).DrawAffineImage(image, OfficeTransform.Translate(1, 0)));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Transparent, image.GetPixel(1, 0));
        }

        [Theory]
        [InlineData("rectangle")]
        [InlineData("projection")]
        [InlineData("affine")]
        [InlineData("blend")]
        public void OverlappingSelfDrawMatchesAnIndependentSourceSnapshot(string route) {
            var image = new OfficeRasterImage(3, 1);
            image.SetPixel(0, 0, OfficeColor.Red);
            image.SetPixel(1, 0, OfficeColor.Green);
            image.SetPixel(2, 0, OfficeColor.Blue);
            var expected = image.Clone();
            var snapshot = image.Clone();
            Draw(new OfficeRasterCanvas(expected), snapshot, route);
            Draw(new OfficeRasterCanvas(image), image, route);
            Assert.Equal(expected.GetPixels(), image.GetPixels());
        }

        private static void Draw(OfficeRasterCanvas canvas, OfficeRasterImage source, string route) {
            switch (route) {
                case "rectangle": canvas.DrawImage(source, 1, 0, 3, 1); break;
                case "projection": canvas.DrawImage(source, new OfficeImageProjection(new OfficeImagePlacement(1, 0, 3, 1))); break;
                case "affine": canvas.DrawAffineImage(source, OfficeTransform.Translate(1, 0)); break;
                case "blend": canvas.DrawAffineImage(source, OfficeTransform.Translate(1, 0), 1, OfficeBlendMode.Screen); break;
                default: throw new ArgumentException(nameof(route));
            }
        }

        [Fact]
        public void PixelWritesClipWhileReadsAndSnapshotsKeepTheirOwnershipContract() {
            byte[] input = { 1, 2, 3, 128 };
            var image = OfficeRasterImage.FromRgba32(1, 1, input);
            input[0] = 99;
            image.SetPixel(-1, 0, OfficeColor.Red);
            image.BlendPixel(1, 0, OfficeColor.Blue);
            byte[] snapshot = image.GetPixels();
            snapshot[0] = 88;
            Assert.Equal(OfficeColor.FromRgba(1, 2, 3, 128), image.GetPixel(0, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => image.GetPixel(-1, 0));
        }

        [Theory]
        [InlineData(OfficeImageFit.Contain, 100, 100, 100, 50, 100, 50, 0)]
        [InlineData(OfficeImageFit.Cover, 100, 100, 100, 100, 200, 100, 50)]
        [InlineData(OfficeImageFit.Stretch, 100, 100, 100, 100, 100, 100, 0)]
        [InlineData(OfficeImageFit.Contain, 100, null, 100, 50, 100, 50, 0)]
        [InlineData(OfficeImageFit.Stretch, 100, null, 100, 200, 100, 200, 0)]
        public void ResizePlanAndPixelsAgreeOnFitAndInferredDimensions(OfficeImageFit fit, int? width, int? height,
            int expectedWidth, int expectedHeight, int resizeWidth, int resizeHeight, int cropX) {
            var source = new OfficeRasterImage(400, 200, OfficeColor.FromRgba(20, 40, 80, 128));
            var options = new OfficeRasterResizeOptions { Width = width, Height = height, Fit = fit, ResamplingMode = OfficeRasterResamplingMode.NearestNeighbor };
            var plan = OfficeRasterResampler.PlanResize(source.Width, source.Height, options);
            var result = OfficeRasterResampler.Resize(source, plan);
            Assert.Equal((expectedWidth, expectedHeight), (plan.Width, plan.Height));
            Assert.Equal((resizeWidth, resizeHeight, cropX), (plan.ResizeWidth, plan.ResizeHeight, plan.CropX));
            Assert.Equal((plan.Width, plan.Height), (result.Width, result.Height));
            Assert.Equal(source.GetPixel(0, 0), result.GetPixel(0, 0));
            result.SetPixel(0, 0, OfficeColor.Red);
            Assert.Equal(OfficeColor.FromRgba(20, 40, 80, 128), source.GetPixel(0, 0));
        }

        [Fact]
        public void ResizePlanCapturesSettingsAndDefinesOddCenteredCrop() {
            var source = new OfficeRasterImage(3, 2, OfficeColor.Blue);
            source.SetPixel(0, 0, OfficeColor.Red);
            source.SetPixel(1, 0, OfficeColor.Green);
            var options = new OfficeRasterResizeOptions { Width = 2, Height = 2, Fit = OfficeImageFit.Cover, ResamplingMode = OfficeRasterResamplingMode.NearestNeighbor };
            var plan = OfficeRasterResampler.PlanResize(3, 2, options);
            options.Width = 50;
            var result = OfficeRasterResampler.Resize(source, plan);
            Assert.Equal((2, 2, 0), (result.Width, result.Height, plan.CropX));
            Assert.Equal(OfficeColor.Red, result.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Green, result.GetPixel(1, 0));
            Assert.Throws<ArgumentException>(() => OfficeRasterResampler.Resize(new OfficeRasterImage(2, 2), plan));
            Assert.Equal(3, OfficeRasterResampler.PlanResize(7, 4, new OfficeRasterResizeOptions { Width = 5 }).Height);
        }

        [Fact]
        public void ResizePlanningRejectsHugeCoverIntermediateAndCropPeakBeforeAllocation() {
            Assert.Throws<ArgumentException>(() => OfficeRasterResampler.PlanResize(10000, 1,
                new OfficeRasterResizeOptions { Width = 1, Height = 10000, Fit = OfficeImageFit.Cover }));
            Assert.Throws<ArgumentException>(() => OfficeRasterResampler.PlanResize(6500, 4000,
                new OfficeRasterResizeOptions { Width = 4000, Height = 4000, Fit = OfficeImageFit.Cover, ResamplingMode = OfficeRasterResamplingMode.NearestNeighbor }));
        }

        [Fact]
        public void FrameResizeKeepsTimingAndRejectsAggregateOutputBeforeMapping() {
            var source = new OfficeRasterImage(1, 1, OfficeColor.Blue);
            var frames = new OfficeRasterFrames(new[] {
                new OfficeRasterFrame(source, TimeSpan.FromMilliseconds(50)), new OfficeRasterFrame(source, TimeSpan.FromMilliseconds(75)), new OfficeRasterFrame(source)
            }, playCount: 4);
            Assert.Throws<ArgumentException>(() => frames.Resize(new OfficeRasterResizeOptions {
                Width = 5000, Height = 5000, ResamplingMode = OfficeRasterResamplingMode.NearestNeighbor
            }));
            var result = frames.Resize(new OfficeRasterResizeOptions { Width = 2, Height = 2 });
            Assert.Equal(4, result.PlayCount);
            Assert.Equal(frames[1].Duration, result[1].Duration);
            Assert.Equal((2, 2), (result[0].Image.Width, result[0].Image.Height));
            result[0].Image.SetPixel(0, 0, OfficeColor.Red);
            Assert.Equal(OfficeColor.Blue, result[1].Image.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Blue, source.GetPixel(0, 0));
            using var canceled = new CancellationTokenSource();
            canceled.Cancel();
            Assert.Throws<OperationCanceledException>(() => frames.Resize(new OfficeRasterResizeOptions { Width = 2 }, canceled.Token));
        }

        [Fact]
        public void PercentageFrameResizePlansVaryingDimensionsAndFloorsBeforeMapping() {
            var frames = new OfficeRasterFrames(new[] {
                new OfficeRasterFrame(new OfficeRasterImage(3, 4, OfficeColor.Blue)),
                new OfficeRasterFrame(new OfficeRasterImage(1, 1, OfficeColor.Red))
            });
            var result = frames.Resize(50);
            Assert.Equal((1, 2), (result[0].Image.Width, result[0].Image.Height));
            Assert.Equal((1, 1), (result[1].Image.Width, result[1].Image.Height));
            Assert.Equal(OfficeColor.Red, result[1].Image.GetPixel(0, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => frames.Resize(0));
            var wide = new OfficeRasterFrames(new[] { new OfficeRasterFrame(new OfficeRasterImage(101, 1)) });
            Assert.Throws<ArgumentOutOfRangeException>(() => wide.Resize(int.MaxValue));
            Assert.Throws<ArgumentException>(() => frames.Resize(50, additionalRetainedBytes: OfficeRasterGuards.MaximumDecodedBytes));
            Assert.Equal((3, 4), (frames[0].Image.Width, frames[0].Image.Height));
            Assert.Equal(OfficeColor.Blue, frames[0].Image.GetPixel(0, 0));
        }

        [Fact]
        public void TextOptionsUseTheExplicitFontForMetricsAndPaint() {
            var fonts = new OfficeFontFaceCollection().Add("Raster contract", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf")));
            var options = new OfficeRasterTextOptions { FontFamily = "Raster contract", Fonts = fonts, FontSize = 18, FontStyle = OfficeFontStyle.Italic, LineHeight = 1.5, Clip = true };
            var actual = new OfficeRasterImage(140, 50, OfficeColor.White);
            var expected = actual.Clone();
            var canvas = new OfficeRasterCanvas(expected, fonts: fonts);
            var layout = OfficeRasterText.Measure("Office", options);
            Assert.Equal(canvas.MeasureText("Office", 18, options.FontFamily, options.FontStyle), layout.Width);
            Assert.Equal(27, layout.LineHeight);
            var plan = OfficeTextBlockRenderPlan.CreateTextBlockFromRectangle("Office", 18, 4, 3, 120, 40,
                (text, size) => canvas.MeasureText(text, size, options.FontFamily, options.FontStyle), OfficeTextAlignment.Left,
                OfficeTextVerticalAlignment.Top, 1.5, 18, false, false, false, OfficeTextOverflowBehavior.Clip);
            using (canvas.PushClipRectangle(4, 3, 120, 40)) {
                OfficeTextBlockRenderer.DrawRasterTextBox(canvas, plan, OfficeColor.Blue, italic: true, fontFamily: options.FontFamily);
            }
            OfficeRasterText.Draw(actual, "Office", 4, 3, 120, 40, OfficeColor.Blue, options);
            Assert.Equal(expected.GetPixels(), actual.GetPixels());
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        public void TextMeasurementRejectsInvalidWrapWidth(double width) {
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterText.Measure("text", new OfficeRasterTextOptions { Wrap = true }, width));
        }

        [Fact]
        public void TextStoragePlanningRejectsRetainedSourceBeforeMappingAndEffectsBeforeMutation() {
            var source = new OfficeRasterImage(30_000, 1_000, OfficeColor.White);
            var frames = new OfficeRasterFrames(new[] { new OfficeRasterFrame(source) });
            var options = new OfficeRasterTextOptions();
            long temporaryBytes = OfficeRasterText.EstimateAdditionalWorkingBytes(source.Width, source.Height, options);
            bool mapped = false;
            Assert.Throws<ArgumentException>(() => frames.Transform(image => {
                mapped = true;
                var result = image.Clone();
                OfficeRasterText.Draw(result, "text", 0, 0, 100, 40, OfficeColor.Black, options);
                return result;
            }, additionalRetainedBytes: temporaryBytes));
            Assert.False(mapped);
            options.OutlineColor = OfficeColor.Red;
            options.OutlineWidth = 1;
            Assert.Throws<ArgumentException>(() => OfficeRasterText.Draw(source, "text", 0, 0, 100, 40, OfficeColor.Black, options));
            Assert.Equal(OfficeColor.White, source.GetPixel(1, 1));
        }

        [Fact]
        public void TextStoragePlanningIncludesEnabledEffectsAndRejectsInvalidInput() {
            var options = new OfficeRasterTextOptions { ShadowColor = OfficeColor.Red, ShadowOffsetX = 3 };
            long maskOnly = OfficeRasterText.EstimateAdditionalWorkingBytes(80, 40, options);
            options.OutlineWidth = 1;
            Assert.Equal(maskOnly, OfficeRasterText.EstimateAdditionalWorkingBytes(80, 40, options));
            options.OutlineColor = OfficeColor.Blue;
            Assert.Equal(maskOnly + 80 * 40 * 4L, OfficeRasterText.EstimateAdditionalWorkingBytes(80, 40, options));
            Assert.Throws<ArgumentNullException>(() => OfficeRasterText.EstimateAdditionalWorkingBytes(80, 40, null!));
            Assert.Throws<ArgumentException>(() => OfficeRasterText.EstimateAdditionalWorkingBytes(50_000_001, 1, options));
            Assert.Throws<ArgumentException>(() => OfficeRasterText.EstimateAdditionalWorkingBytes(0, 40, options));
            options.FontSize = double.NaN;
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterText.EstimateAdditionalWorkingBytes(80, 40, options));
        }

        [Fact]
        public void PlannedSmallFrameTextEffectsStillPaintIndependentResults() {
            var fonts = new OfficeFontFaceCollection().Add("Raster contract", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf")));
            var options = new OfficeRasterTextOptions { Fonts = fonts, FontFamily = "Raster contract", FontSize = 14, Clip = true,
                ShadowColor = OfficeColor.Red, ShadowOffsetX = 2, ShadowOffsetY = 2, OutlineColor = OfficeColor.Blue, OutlineWidth = 1 };
            options.TextShapingProvider = new OriginalOptionsMutationProvider(() => { options.FontSize = double.NaN; options.OutlineWidth = 65; });
            OfficeRasterTextOptions captured = options.Clone();
            Assert.NotSame(options.Fonts, captured.Fonts);
            var first = new OfficeRasterImage(80, 40, OfficeColor.White);
            var second = new OfficeRasterImage(60, 30, OfficeColor.White);
            var frames = new OfficeRasterFrames(new[] { new OfficeRasterFrame(first), new OfficeRasterFrame(second) });
            long temporaryBytes = Math.Max(OfficeRasterText.EstimateAdditionalWorkingBytes(first.Width, first.Height, captured),
                OfficeRasterText.EstimateAdditionalWorkingBytes(second.Width, second.Height, captured));
            var result = frames.Transform(source => {
                var painted = source.Clone();
                OfficeRasterText.Draw(painted, "Text", 2, 2, 50, 25, OfficeColor.Black, captured);
                return painted;
            }, additionalRetainedBytes: temporaryBytes);
            for (int i = 0; i < result.Count; i++) {
                Assert.NotEqual(frames[i].Image.GetPixels(), result[i].Image.GetPixels());
            }
            Assert.Equal(OfficeColor.White, first.GetPixel(5, 5));
            Assert.Equal(OfficeColor.White, second.GetPixel(5, 5));
            Assert.True(double.IsNaN(options.FontSize));
            Assert.Equal(14, captured.FontSize);
            Assert.Equal(1, captured.OutlineWidth);
        }

        private sealed class OriginalOptionsMutationProvider : IOfficeTextShapingProvider {
            private readonly Action _mutateOriginal;
            internal OriginalOptionsMutationProvider(Action mutateOriginal) => _mutateOriginal = mutateOriginal;
            public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
                _mutateOriginal();
                return null;
            }
        }
    }
}
