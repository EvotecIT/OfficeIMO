using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class DrawingRasterTests {
        [Fact]
        public void OfficeRasterImageDecoder_DecodesGifFirstFrameThroughSharedRasterPath() {
            byte[] gif = CreateIndexedGif(
                2,
                2,
                new[] { OfficeColor.Red, OfficeColor.Lime, OfficeColor.Blue, OfficeColor.White },
                new byte[] { 0, 1, 2, 3 });

            Assert.True(OfficeRasterImageDecoder.TryDecode(gif, out OfficeRasterImage? image));
            Assert.Equal(2, image!.Width);
            Assert.Equal(2, image.Height);
            Assert.Equal(OfficeColor.Red, image.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Lime, image.GetPixel(1, 0));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(0, 1));
            Assert.Equal(OfficeColor.White, image.GetPixel(1, 1));
        }

        [Fact]
        public void OfficeRasterImageDecoder_DecodesInterlacedGifRowsThroughSharedRasterPath() {
            byte[] gif = CreateIndexedGif(
                1,
                4,
                new[] { OfficeColor.Red, OfficeColor.Lime, OfficeColor.Blue, OfficeColor.White },
                new byte[] { 0, 1, 2, 3 },
                interlaced: true);

            Assert.True(OfficeRasterImageDecoder.TryDecode(gif, out OfficeRasterImage? image));
            Assert.Equal(1, image!.Width);
            Assert.Equal(4, image.Height);
            Assert.Equal(OfficeColor.Red, image.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Lime, image.GetPixel(0, 1));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(0, 2));
            Assert.Equal(OfficeColor.White, image.GetPixel(0, 3));
        }

        [Fact]
        public void OfficeDrawingRasterRenderer_PaintsDecodedGifImages() {
            byte[] gif = CreateSinglePixelGif();
            OfficeDrawing drawing = new OfficeDrawing(20, 16);
            drawing.AddImage(
                gif,
                "image/gif",
                new OfficeImageProjection(new OfficeImagePlacement(4, 3, 8, 6)),
                "GIF marker");

            OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.Black);

            Assert.Equal(OfficeColor.White, rendered.GetPixel(7, 5));
        }

        private static byte[] CreateSinglePixelGif() =>
            Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");

        private static byte[] CreateIndexedGif(int width, int height, IReadOnlyList<OfficeColor> palette, IReadOnlyList<byte> pixels, bool interlaced = false) {
            if (pixels.Count != width * height) {
                throw new ArgumentException("Pixel count must match GIF dimensions.", nameof(pixels));
            }

            int colorTableSize = 2;
            while (colorTableSize < palette.Count) {
                colorTableSize *= 2;
            }

            int minimumCodeSize = Math.Max(2, GetRequiredBits(colorTableSize - 1));
            byte[] lzw = EncodeGifLzw(ReorderGifPixels(width, height, pixels, interlaced), minimumCodeSize);
            var bytes = new List<byte>();
            bytes.AddRange(new byte[] { (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a' });
            WriteUInt16LittleEndian(bytes, width);
            WriteUInt16LittleEndian(bytes, height);
            bytes.Add((byte)(0x80 | ((minimumCodeSize - 1) << 4) | (GetRequiredBits(colorTableSize - 1) - 1)));
            bytes.Add(0);
            bytes.Add(0);
            for (int i = 0; i < colorTableSize; i++) {
                OfficeColor color = i < palette.Count ? palette[i] : OfficeColor.Black;
                bytes.Add(color.R);
                bytes.Add(color.G);
                bytes.Add(color.B);
            }

            bytes.Add(0x2C);
            WriteUInt16LittleEndian(bytes, 0);
            WriteUInt16LittleEndian(bytes, 0);
            WriteUInt16LittleEndian(bytes, width);
            WriteUInt16LittleEndian(bytes, height);
            bytes.Add(interlaced ? (byte)0x40 : (byte)0x00);
            bytes.Add((byte)minimumCodeSize);
            bytes.Add((byte)lzw.Length);
            bytes.AddRange(lzw);
            bytes.Add(0);
            bytes.Add(0x3B);
            return bytes.ToArray();
        }

        private static byte[] ReorderGifPixels(int width, int height, IReadOnlyList<byte> pixels, bool interlaced) {
            if (!interlaced) {
                return pixels.ToArray();
            }

            var reordered = new List<byte>(pixels.Count);
            foreach (int y in EnumerateGifRows(height)) {
                for (int x = 0; x < width; x++) {
                    reordered.Add(pixels[(y * width) + x]);
                }
            }

            return reordered.ToArray();
        }

        private static byte[] EncodeGifLzw(IReadOnlyList<byte> indices, int minimumCodeSize) {
            int clearCode = 1 << minimumCodeSize;
            int endCode = clearCode + 1;
            int dictionaryCount = clearCode + 2;
            int codeSize = minimumCodeSize + 1;
            int previousCode = -1;
            var bits = new List<int>();

            WriteBits(bits, clearCode, codeSize);
            for (int i = 0; i < indices.Count; i++) {
                WriteBits(bits, indices[i], codeSize);
                if (previousCode >= 0 && dictionaryCount < 4096) {
                    dictionaryCount++;
                    if (dictionaryCount == (1 << codeSize) && codeSize < 12) {
                        codeSize++;
                    }
                }

                previousCode = indices[i];
            }

            WriteBits(bits, endCode, codeSize);
            var bytes = new byte[(bits.Count + 7) / 8];
            for (int i = 0; i < bits.Count; i++) {
                bytes[i / 8] |= (byte)(bits[i] << (i % 8));
            }

            return bytes;
        }

        private static void WriteBits(List<int> bits, int value, int count) {
            for (int i = 0; i < count; i++) {
                bits.Add((value >> i) & 1);
            }
        }

        private static IEnumerable<int> EnumerateGifRows(int height) {
            int[] starts = { 0, 4, 2, 1 };
            int[] steps = { 8, 8, 4, 2 };
            for (int pass = 0; pass < starts.Length; pass++) {
                for (int y = starts[pass]; y < height; y += steps[pass]) {
                    yield return y;
                }
            }
        }

        private static int GetRequiredBits(int value) {
            int bits = 0;
            do {
                bits++;
                value >>= 1;
            } while (value > 0);

            return bits;
        }

        private static void WriteUInt16LittleEndian(List<byte> bytes, int value) {
            bytes.Add((byte)(value & 0xFF));
            bytes.Add((byte)((value >> 8) & 0xFF));
        }
    }
}
