using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public class DrawingPngWriterTests {
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
        public void OfficePngWriter_RejectsIndexedColorWithoutPalette() {
            byte[] scanlines = { 0, 0 };

            Assert.Throws<ArgumentOutOfRangeException>(() => OfficePngWriter.EncodeScanlines(1, 1, 8, 3, scanlines));
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficePngWriter.CreateFromCompressedScanlines(1, 1, 8, 3, Array.Empty<byte>()));
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
