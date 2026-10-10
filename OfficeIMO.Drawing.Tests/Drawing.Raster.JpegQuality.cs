using OfficeIMO.Drawing;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class RasterJpegQualityTests {
        [Theory]
        [InlineData(OfficeJpegSubsampling.Y420, false)]
        [InlineData(OfficeJpegSubsampling.Y420, true)]
        [InlineData(OfficeJpegSubsampling.Y422, false)]
        [InlineData(OfficeJpegSubsampling.Y422, true)]
        public void CommonDecodeInterpolatesSubsampledChromaByDefault(OfficeJpegSubsampling subsampling, bool progressive) {
            byte[] jpeg = EncodeChromaEdges(subsampling, progressive);
            OfficeRasterImage nearest = OfficeJpegCodec.Decode(jpeg, new OfficeJpegDecodeOptions(highQualityChroma: false));
            OfficeRasterImage interpolated = OfficeJpegCodec.Decode(jpeg, new OfficeJpegDecodeOptions(highQualityChroma: true));
            Assert.False(nearest.GetPixels().SequenceEqual(interpolated.GetPixels()));

            Assert.True(OfficeRasterImageDecoder.TryDecode(jpeg, out OfficeRasterImage? decoded));
            Assert.Equal(interpolated.GetPixels(), decoded!.GetPixels());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void SelectedChromaPolicySurvivesSnapshotsStreamsAndSequences(bool highQualityChroma) {
            byte[] jpeg = EncodeChromaEdges(OfficeJpegSubsampling.Y420, progressive: false);
            byte[] expected = OfficeJpegCodec.Decode(jpeg,
                new OfficeJpegDecodeOptions(highQualityChroma: highQualityChroma)).GetPixels();
            var options = new OfficeRasterDecodeOptions { JpegHighQualityChroma = highQualityChroma };
            OfficeRasterDecodeOptions snapshot = options.Clone();
            options.JpegHighQualityChroma = !highQualityChroma;

            Assert.Equal(expected, OfficeRasterImageDecoder.Decode(jpeg, snapshot).GetPixels());
            Assert.Equal(expected, Assert.Single(OfficeRasterImageDecoder.DecodeFrames(jpeg, snapshot)).Image.GetPixels());

            var prefixed = new byte[jpeg.Length + 7];
            Buffer.BlockCopy(jpeg, 0, prefixed, 7, jpeg.Length);
            using var stream = new MemoryStream(prefixed);
            stream.Position = 7;
            Assert.Equal(expected, OfficeRasterImageDecoder.Decode(stream, snapshot).GetPixels());
            Assert.Equal(7, stream.Position);
            Assert.Equal(expected, Assert.Single(OfficeRasterImageDecoder.DecodeFrames(stream, snapshot)).Image.GetPixels());
            Assert.Equal(7, stream.Position);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ChromaQualityDoesNotPermitIncompleteJpegPayloads(bool highQualityChroma) {
            byte[] jpeg = EncodeChromaEdges(OfficeJpegSubsampling.Y420, progressive: false);
            Array.Resize(ref jpeg, jpeg.Length - 2);

            Assert.False(OfficeRasterImageDecoder.TryDecode(jpeg,
                new OfficeRasterDecodeOptions { JpegHighQualityChroma = highQualityChroma }, out OfficeRasterImage? image, out _));
            Assert.Null(image);
        }

        private static byte[] EncodeChromaEdges(OfficeJpegSubsampling subsampling, bool progressive) {
            var source = new OfficeRasterImage(31, 19);
            for (int y = 0; y < source.Height; y++) {
                for (int x = 0; x < source.Width; x++) {
                    source.SetPixel(x, y, (x + y) % 12 < 6
                        ? OfficeColor.FromRgb(240, 20, 160)
                        : OfficeColor.FromRgb(15, 190, 235));
                }
            }
            return OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Quality = 90, Subsampling = subsampling, Progressive = progressive
            });
        }
    }
}