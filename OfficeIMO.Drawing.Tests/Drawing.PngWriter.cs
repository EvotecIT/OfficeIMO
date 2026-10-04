using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public class DrawingPngWriterTests {
        [Theory]
        [InlineData(257, 129, OfficePngCompression.Optimal)]
        [InlineData(257, 129, OfficePngCompression.Stored)]
        [InlineData(16381, 1, OfficePngCompression.Stored)]
        public void MaterializedPngEncodersPreserveLargeRgbaAndResolution(int width, int height, OfficePngCompression compression) {
            var image = new OfficeRasterImage(width, height);
            uint state = 0x915DF32B;
            for (int y = 0; y < height; y++) {
                for (int x = 0; x < width; x++) {
                    state ^= state << 13;
                    state ^= state >> 17;
                    state ^= state << 5;
                    image.SetPixel(x, y, OfficeColor.FromRgba(
                        (byte)state, (byte)(state >> 8), (byte)(state >> 16), (byte)(state >> 24)));
                }
            }

            byte[] expected = image.GetPixels();
            var options = new OfficePngEncodeOptions { Compression = compression, DpiX = 144, DpiY = 120 };
            var withoutMetadata = new OfficePngEncodeOptions {
                Compression = compression, DpiX = 144, DpiY = 120, WritePhysicalResolution = false
            };
            byte[][] withResolution = {
                OfficePngWriter.Encode(image, options),
                OfficePngWriter.EncodeRgba(width, height, expected, options)
            };
            byte[][] withoutResolution = {
                OfficePngWriter.Encode(image, compression),
                OfficePngWriter.EncodeRgba(width, height, expected, compression),
                OfficePngWriter.Encode(image, withoutMetadata),
                OfficePngWriter.EncodeRgba(width, height, expected, withoutMetadata)
            };

            foreach (byte[] png in withResolution.Concat(withoutResolution)) {
                Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? decoded));
                Assert.NotNull(decoded);
                Assert.Equal(expected, decoded!.GetPixels());
            }
            foreach (byte[] png in withResolution) {
                OfficeImageInfo info = OfficeImageReader.Identify(png);
                Assert.InRange(info.DpiX, 143.98, 144.02);
                Assert.InRange(info.DpiY, 119.98, 120.02);
            }
            foreach (byte[] png in withoutResolution) {
                Assert.Throws<InvalidOperationException>(() => ExtractChunk(png, "pHYs"));
            }
        }

        [Theory]
        [InlineData(1)]
        [InlineData(5)]
        [InlineData(6)]
        [InlineData(7)]
        [InlineData(4095)]
        [InlineData(4097)]
        public void TrueColorScanlinesExpandEveryRgbChannelAndOpaqueAlphaAcrossTails(int width) {
            const int height = 3;
            var scanlines = new byte[(width * 3 + 1) * height];
            var expected = new byte[width * height * 4];
            for (int y = 0; y < height; y++) {
                for (int x = 0; x < width; x++) {
                    int source = y * (width * 3 + 1) + 1 + x * 3;
                    int target = (y * width + x) * 4;
                    for (int channel = 0; channel < 3; channel++) {
                        byte value = unchecked((byte)(x * (31 + channel * 46) + y * 128 + channel));
                        scanlines[source + channel] = value;
                        expected[target + channel] = value;
                    }
                    expected[target + 3] = 255;
                }
            }

            byte[] png = OfficePngWriter.EncodeScanlines(width, height, 8, 2, scanlines);

            Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? decoded));
            Assert.NotNull(decoded);
            Assert.Equal(expected, decoded!.GetPixels());
        }

        [Theory]
        [InlineData(15)]
        [InlineData(16)]
        [InlineData(17)]
        [InlineData(1023)]
        [InlineData(1024)]
        [InlineData(1025)]
        public void AdaptiveFilteringPreservesRgbaAcrossArithmeticAndBlockTails(int width) {
            var image = new OfficeRasterImage(width, 3);
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < width; x++) {
                    image.SetPixel(x, y, OfficeColor.FromRgba(
                        unchecked((byte)(x * 127 + y * 128)),
                        unchecked((byte)(x * 31 - y * 128)),
                        unchecked((byte)(x * 255 + y)),
                        unchecked((byte)(x * 17 + y * 128))));
                }
            }

            byte[] png = OfficePngWriter.Encode(image);
            using var stream = new MemoryStream();
            OfficePngWriter.EncodeTo(image, stream);

            foreach (byte[] encoded in new[] { png, stream.ToArray() }) {
                Assert.True(OfficePngReader.TryDecode(encoded, out OfficeRasterImage? decoded));
                Assert.NotNull(decoded);
                Assert.Equal(image.GetPixels(), decoded!.GetPixels());
            }
        }

        [Fact]
        public void OfficePngWriter_EncodesSharedPngScanlineContainers() {
            byte[] scanlines = { 0, 255, 0, 0, 128 };

            byte[] png = OfficePngWriter.EncodeScanlines(1, 1, 8, 6, scanlines, OfficePngCompression.Stored);
            byte[] wrapped = OfficePngWriter.CreateFromCompressedScanlines(1, 1, 8, 6, ExtractChunk(png, "IDAT"));

            Assert.Equal(6, png[25]);
            Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? decoded));
            Assert.NotNull(decoded);
            Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 128), decoded!.GetPixel(0, 0));
            Assert.True(OfficePngReader.TryDecode(wrapped, out OfficeRasterImage? wrappedDecoded));
            Assert.NotNull(wrappedDecoded);
            Assert.Equal(decoded.GetPixel(0, 0), wrappedDecoded!.GetPixel(0, 0));
        }

        [Fact]
        public void OfficePngWriter_CanEncodeRasterImagesWithStoredCompression() {
            OfficeRasterImage image = new OfficeRasterImage(1, 1, OfficeColor.Transparent);
            image.SetPixel(0, 0, OfficeColor.FromRgba(255, 0, 0, 128));

            byte[] png = OfficePngWriter.Encode(image, OfficePngCompression.Stored);
            byte[] idat = ExtractChunk(png, "IDAT");

            Assert.Equal(0x78, idat[0]);
            Assert.Equal(0x01, idat[1]);
            Assert.Equal(1, idat[2]);
            Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? decoded));
            Assert.NotNull(decoded);
            Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 128), decoded!.GetPixel(0, 0));
        }

        [Fact]
        public void OfficePngWriter_AdaptiveFilteringPreservesPixelsAndCompressesStructuredRows() {
            var image = new OfficeRasterImage(128, 64);
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    image.SetPixel(
                        x,
                        y,
                        OfficeColor.FromRgba((byte)x, (byte)(x + y), (byte)(y * 3), (byte)(128 + (x & 127))));
                }
            }

            byte[] optimal = OfficePngWriter.Encode(image, OfficePngCompression.Optimal);
            byte[] stored = OfficePngWriter.Encode(image, OfficePngCompression.Stored);

            Assert.True(OfficePngReader.TryDecode(optimal, out OfficeRasterImage? decoded));
            Assert.NotNull(decoded);
            Assert.Equal(image.GetPixels(), decoded!.GetPixels());
            Assert.True(optimal.Length < stored.Length / 4, $"Expected adaptive PNG ({optimal.Length}) to be materially smaller than stored PNG ({stored.Length}).");
        }

        [Fact]
        public void OfficePngWriter_RejectsIndexedColorWithoutPalette() {
            byte[] scanlines = { 0, 0 };

            Assert.Throws<ArgumentOutOfRangeException>(() => OfficePngWriter.EncodeScanlines(1, 1, 8, 3, scanlines));
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficePngWriter.CreateFromCompressedScanlines(1, 1, 8, 3, Array.Empty<byte>()));
        }

        [Fact]
        public void OfficePngWriter_RejectsInvalidColorTypeBitDepthPairs() {
            byte[] scanlines = { 0, 0, 0, 0 };

            Assert.Throws<ArgumentOutOfRangeException>(() => OfficePngWriter.EncodeScanlines(1, 1, 4, 2, scanlines));
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficePngWriter.EncodeScanlines(1, 1, 4, 4, scanlines));
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficePngWriter.CreateFromCompressedScanlines(1, 1, 4, 6, Array.Empty<byte>()));
        }

        [Fact]
        public void OfficePngWriter_RejectsScanlineBuffersThatDoNotMatchIhdrLayout() {
            byte[] shortRgbaScanline = { 0, 255, 0, 0, 255 };
            byte[] shortGrayscaleScanlines = { 0, 0, 0 };

            Assert.Throws<ArgumentException>(() => OfficePngWriter.EncodeScanlines(2, 1, 8, 6, shortRgbaScanline));
            Assert.Throws<ArgumentException>(() => OfficePngWriter.EncodeScanlines(8, 2, 1, 0, shortGrayscaleScanlines));
        }

        private static byte[] ExtractChunk(byte[] png, string type) {
            int offset = 8;
            while (offset + 8 <= png.Length) {
                int length = (png[offset] << 24) |
                    (png[offset + 1] << 16) |
                    (png[offset + 2] << 8) |
                    png[offset + 3];
                string currentType = System.Text.Encoding.ASCII.GetString(png, offset + 4, 4);
                int dataOffset = offset + 8;
                if (currentType == type) {
                    byte[] data = new byte[length];
                    Buffer.BlockCopy(png, dataOffset, data, 0, length);
                    return data;
                }

                offset = dataOffset + length + 4;
            }

            throw new InvalidOperationException("PNG chunk was not found.");
        }
    }
}
