using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public class DrawingImageOptimizationTests {
        [Fact]
        public void OfficeRasterResampler_NearestNeighborPreservesSourceQuadrants() {
            var source = new OfficeRasterImage(2, 2);
            source.SetPixel(0, 0, OfficeColor.Red);
            source.SetPixel(1, 0, OfficeColor.Lime);
            source.SetPixel(0, 1, OfficeColor.Blue);
            source.SetPixel(1, 1, OfficeColor.White);

            OfficeRasterImage resized = OfficeRasterResampler.Resize(source, 4, 4, OfficeRasterResamplingMode.NearestNeighbor);

            Assert.Equal(OfficeColor.Red, resized.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Lime, resized.GetPixel(3, 0));
            Assert.Equal(OfficeColor.Blue, resized.GetPixel(0, 3));
            Assert.Equal(OfficeColor.White, resized.GetPixel(3, 3));
        }

        [Fact]
        public void OfficeRasterResampler_BilinearUsesPremultipliedAlpha() {
            var source = new OfficeRasterImage(2, 1);
            source.SetPixel(0, 0, OfficeColor.FromRgba(0, 0, 255, 0));
            source.SetPixel(1, 0, OfficeColor.Red);

            OfficeColor blended = OfficeRasterResampler.Resize(source, 1, 1).GetPixel(0, 0);

            Assert.InRange(blended.R, 254, 255);
            Assert.InRange(blended.G, 0, 1);
            Assert.InRange(blended.B, 0, 1);
            Assert.InRange(blended.A, 127, 128);
        }

        [Fact]
        public void OfficeJpegCodec_RoundTripsDimensionsAndRepresentativeColors() {
            OfficeRasterImage source = CreateQuadrantImage(32, 24);

            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Quality = 92,
                Subsampling = OfficeJpegSubsampling.Y444
            });

            Assert.True(OfficeJpegCodec.IsJpeg(jpeg));
            Assert.True(OfficeRasterImageDecoder.TryDecode(jpeg, out OfficeRasterImage? decoded));
            Assert.NotNull(decoded);
            Assert.Equal(32, decoded!.Width);
            Assert.Equal(24, decoded.Height);
            AssertColorNear(decoded.GetPixel(4, 4), OfficeColor.Red, 20);
            AssertColorNear(decoded.GetPixel(27, 4), OfficeColor.Lime, 20);
            AssertColorNear(decoded.GetPixel(4, 19), OfficeColor.Blue, 20);
            AssertColorNear(decoded.GetPixel(27, 19), OfficeColor.White, 20);
        }

        [Theory]
        [InlineData(OfficeJpegSubsampling.Y444, false, false)]
        [InlineData(OfficeJpegSubsampling.Y422, false, true)]
        [InlineData(OfficeJpegSubsampling.Y420, true, false)]
        [InlineData(OfficeJpegSubsampling.Y420, true, true)]
        public void OfficeJpegCodec_EncodesManagedVariants(OfficeJpegSubsampling subsampling, bool progressive, bool optimizeHuffman) {
            OfficeRasterImage source = CreateQuadrantImage(37, 29);

            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Quality = 88,
                Subsampling = subsampling,
                Progressive = progressive,
                OptimizeHuffman = optimizeHuffman
            });
            OfficeRasterImage decoded = OfficeJpegCodec.Decode(jpeg, new OfficeJpegDecodeOptions(highQualityChroma: true));

            Assert.Equal(source.Width, decoded.Width);
            Assert.Equal(source.Height, decoded.Height);
            AssertColorNear(decoded.GetPixel(4, 4), OfficeColor.Red, 28);
            AssertColorNear(decoded.GetPixel(32, 24), OfficeColor.White, 28);
        }

        [Fact]
        public void OfficeJpegMetadata_SnapshotsCallerBuffers() {
            byte[] exif = { 1, 2, 3 };
            var metadata = new OfficeJpegMetadata(exif: exif);

            exif[0] = 9;
            byte[] exposed = metadata.Exif!;
            exposed[1] = 8;

            Assert.Equal(new byte[] { 1, 2, 3 }, metadata.Exif);
        }

        [Fact]
        public void OfficeJpegCodec_DecodesIndependentJpegFixture() {
            byte[] jpeg = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Kulek.jpg"));
            OfficeImageInfo identified = OfficeImageReader.Identify(jpeg);

            OfficeRasterImage decoded = OfficeJpegCodec.Decode(jpeg);
            Assert.Equal(OfficeImageFormat.Jpeg, identified.Format);
            Assert.Equal(identified.Width, decoded.Width);
            Assert.Equal(identified.Height, decoded.Height);
            Assert.True(decoded.Width > 100);
            Assert.True(decoded.Height > 100);
        }

        [Fact]
        public void OfficeJpegCodec_FlattensTransparencyAgainstConfiguredBackground() {
            var source = new OfficeRasterImage(8, 8, OfficeColor.FromRgba(255, 0, 0, 128));

            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Quality = 100,
                Background = OfficeColor.Blue
            });
            OfficeRasterImage decoded = OfficeJpegCodec.Decode(jpeg);

            AssertColorNear(decoded.GetPixel(4, 4), OfficeColor.FromRgb(128, 0, 127), 12);
        }

        [Fact]
        public void OfficeImageOptimizer_DownsamplesUsingPlacementBounds() {
            OfficeRasterImage source = CreateQuadrantImage(160, 120);
            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions { Quality = 94 });
            var request = new OfficeImageOptimizationRequest(40, 40) {
                JpegQuality = 80,
                KeepOriginalWhenNotSmaller = false
            };

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(jpeg, request);

            Assert.True(result.Changed);
            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal(160, result.Original.Width);
            Assert.Equal(120, result.Original.Height);
            Assert.Equal(40, result.Final.Width);
            Assert.Equal(30, result.Final.Height);
            Assert.Equal(OfficeImageFormat.Jpeg, result.Final.Format);
            Assert.Equal(jpeg.LongLength, result.OriginalEncodedLength);
            Assert.Equal(result.Bytes.LongLength, result.FinalEncodedLength);
            Assert.True(OfficeJpegCodec.TryDecode(result.Bytes, out OfficeRasterImage? decoded));
            Assert.Equal(40, decoded!.Width);
            Assert.Equal(30, decoded.Height);
        }

        [Fact]
        public void OfficeImageOptimizer_DownsamplesPngWithoutLosingAlpha() {
            var source = new OfficeRasterImage(80, 40, OfficeColor.FromRgba(20, 80, 200, 96));
            byte[] png = OfficePngWriter.Encode(source);

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                png,
                new OfficeImageOptimizationRequest(20, 20) {
                    KeepOriginalWhenNotSmaller = false
                });

            Assert.True(result.Changed);
            Assert.Equal(OfficeImageFormat.Png, result.Final.Format);
            Assert.Equal(20, result.Final.Width);
            Assert.Equal(10, result.Final.Height);
            Assert.True(OfficeRasterImageDecoder.TryDecode(result.Bytes, out OfficeRasterImage? decoded));
            Assert.InRange(decoded!.GetPixel(10, 5).A, 95, 97);
        }

        [Fact]
        public void OfficeImageOptimizer_DoesNotRewriteAnimationCapableFormats() {
            byte[] gif = {
                (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a',
                1, 0, 1, 0, 0, 0, 0
            };

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(gif, new OfficeImageOptimizationRequest(1, 1));

            Assert.Equal(OfficeImageOptimizationStatus.UnsupportedFormat, result.Status);
            Assert.False(result.Changed);
            Assert.Equal(gif, result.Bytes);
        }

        private static OfficeRasterImage CreateQuadrantImage(int width, int height) {
            var image = new OfficeRasterImage(width, height);
            for (int y = 0; y < height; y++) {
                for (int x = 0; x < width; x++) {
                    OfficeColor color = x < width / 2
                        ? (y < height / 2 ? OfficeColor.Red : OfficeColor.Blue)
                        : (y < height / 2 ? OfficeColor.Lime : OfficeColor.White);
                    image.SetPixel(x, y, color);
                }
            }
            return image;
        }

        private static void AssertColorNear(OfficeColor actual, OfficeColor expected, int tolerance) {
            Assert.InRange((int)actual.R, Math.Max(0, expected.R - tolerance), Math.Min(255, expected.R + tolerance));
            Assert.InRange((int)actual.G, Math.Max(0, expected.G - tolerance), Math.Min(255, expected.G + tolerance));
            Assert.InRange((int)actual.B, Math.Max(0, expected.B - tolerance), Math.Min(255, expected.B + tolerance));
            Assert.Equal(255, actual.A);
        }
    }
}
