using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingJpegColorConversionTests {
    [Theory]
    [InlineData(OfficeJpegSubsampling.Y444, false)]
    [InlineData(OfficeJpegSubsampling.Y422, false)]
    [InlineData(OfficeJpegSubsampling.Y420, false)]
    [InlineData(OfficeJpegSubsampling.Y444, true)]
    [InlineData(OfficeJpegSubsampling.Y422, true)]
    [InlineData(OfficeJpegSubsampling.Y420, true)]
    public void BlockAlignedColorTilesDecodeIdenticallyAcrossSmallAndLargeImages(OfficeJpegSubsampling sampling, bool progressive) {
        OfficeRasterImage small = CreateTiles(32, 32);
        OfficeRasterImage large = CreateTiles(512, 513);
        var options = new OfficeRasterEncodingOptions { WriteResolutionMetadata = false,
            Jpeg = new OfficeJpegEncodeOptions { Quality = 85, Subsampling = sampling, Progressive = progressive } };
        byte[] reference = OfficeJpegCodec.Decode(OfficeRasterImageEncoder.Encode(small, OfficeImageExportFormat.Jpeg, options)).GetPixels();
        byte[] encoded = OfficeRasterImageEncoder.Encode(large, OfficeImageExportFormat.Jpeg, options);
        byte[] actual = OfficeJpegCodec.Decode(encoded).GetPixels();
        var expected = new byte[large.Width * large.Height * 4];
        for (int y = 0; y < large.Height; y++) {
            for (int x = 0; x < large.Width; x++) {
                Buffer.BlockCopy(reference, ((y % 32) * 32 + x % 32) * 4, expected, (y * large.Width + x) * 4, 4);
            }
        }
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void PackedNearestRgbMatchesRgbaForOddWidthSubsampledJpeg() {
        var source = new OfficeRasterImage(35, 19);
        for (int y = 0; y < source.Height; y++) {
            for (int x = 0; x < source.Width; x++) {
                source.SetPixel(x, y, OfficeColor.FromRgb(
                    (byte)(x * 6 + y), (byte)(y * 11 + x), (byte)(235 - x * 4 - y * 3)));
            }
        }
        byte[] jpeg = OfficeJpegCodec.Encode(source,
            new OfficeJpegEncodeOptions { Quality = 85, Subsampling = OfficeJpegSubsampling.Y420 });
        byte[] rgba = OfficeJpegCodec.Decode(jpeg).GetPixels();
        var expected = new byte[source.Width * source.Height * 3];
        for (int pixel = 0; pixel < source.Width * source.Height; pixel++) {
            Buffer.BlockCopy(rgba, pixel * 4, expected, pixel * 3, 3);
        }

        Assert.True(OfficeJpegCodec.TryDecodeColorComponents(
            jpeg, requestedColorTransform: null, usePdfColorTransformDefault: true,
            out byte[] rgb, out int width, out int height, out int components));
        Assert.Equal((source.Width, source.Height, 3), (width, height, components));
        Assert.Equal(expected, rgb);
    }

    // Smooth chroma converts whole rows at a time; the per-pixel conversion, used when
    // the decode budget leaves no room for the row tables, is the reference. The noisy
    // content puts interpolated chroma on every sixteenth and many sums near a rounding
    // tie, and the odd size leaves partial blocks and row tails.
    [Theory]
    [InlineData(OfficeJpegSubsampling.Y444)]
    [InlineData(OfficeJpegSubsampling.Y422)]
    [InlineData(OfficeJpegSubsampling.Y420)]
    public void SmoothChromaRowsMatchPerPixelConversion(OfficeJpegSubsampling sampling) {
        var source = new OfficeRasterImage(203, 77);
        var random = new Random(4021);
        for (int y = 0; y < source.Height; y++) {
            for (int x = 0; x < source.Width; x++) {
                source.SetPixel(x, y, OfficeColor.FromRgb((byte)random.Next(256), (byte)(x + random.Next(48)), (byte)(y * 3 + random.Next(32))));
            }
        }
        byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions { Quality = 95, Subsampling = sampling });
        var options = new OfficeJpegDecodeOptions(highQualityChroma: true);
        byte[] rows = OfficeJpegCodec.Decode(jpeg, options).GetPixels();
        Assert.Equal(DecodeWithoutRowTables(jpeg, options), rows);
    }

    // Green fixed-point sums for these (luma, Cb, Cr) sixteenths fall within a unit of a
    // rounding tie, so the row path must fall back to the double expression. A 4:2:0
    // image of flat blocks (quantization 1, DC only) places each triple exactly: at the
    // MCU corner (16j, 16) the interpolated chroma is 9a + 3b + 3c + d of the four
    // surrounding chroma blocks.
    [Fact]
    public void SmoothChromaNearRoundingTiesMatchPerPixelConversion() {
        (int Luma, int Cb, int Cr)[] targets = { (129, 2322, 1143), (113, 3749, 2651), (190, 722, 2743), (144, 2322, 1143) };
        const int mcuColumns = 8, mcuRows = 2;
        var luma = new int[mcuRows * 2, mcuColumns * 2];
        var cb = new int[mcuRows, mcuColumns];
        var cr = new int[mcuRows, mcuColumns];
        for (int row = 0; row < luma.GetLength(0); row++) for (int column = 0; column < luma.GetLength(1); column++) luma[row, column] = 128;
        for (int target = 0; target < targets.Length; target++) {
            int column = target * 2 + 1;
            luma[2, column * 2] = targets[target].Luma;
            PlaceCornerChroma(cb, column, targets[target].Cb);
            PlaceCornerChroma(cr, column, targets[target].Cr);
        }
        byte[] jpeg = WriteFlatBlockJpeg420(luma, cb, cr);
        var options = new OfficeJpegDecodeOptions(highQualityChroma: true);
        byte[] rows = OfficeJpegCodec.Decode(jpeg, options).GetPixels();
        Assert.Equal(DecodeWithoutRowTables(jpeg, options), rows);
    }

    // Chroma blocks (0, j - 1), (0, j), (1, j - 1), (1, j) weigh 1, 3, 3 and 9 at (16j, 16).
    private static void PlaceCornerChroma(int[,] chroma, int column, int sixteenths) {
        int value = Math.Min(sixteenths / 16, 250);
        int remainder = sixteenths - value * 16;
        chroma[1, column] = value;
        chroma[1, column - 1] = chroma[0, column] = value + remainder / 6;
        chroma[0, column - 1] = value + remainder % 6;
    }

    private static byte[] DecodeWithoutRowTables(byte[] jpeg, OfficeJpegDecodeOptions options) {
        // The largest retained size that still decodes leaves no budget for row tables.
        long low = 0L;
        long high = OfficeRasterGuards.MaximumDecodedBytes;
        while (low < high) {
            long middle = low + (high - low + 1L) / 2L;
            if (OfficeJpegCodec.TryDecode(jpeg, CancellationToken.None, middle, out _, options)) low = middle;
            else high = middle - 1L;
        }
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, CancellationToken.None, low, out OfficeRasterImage? image, options));
        return image!.GetPixels();
    }

    // Baseline 4:2:0 JPEG whose blocks hold only DC, with every quantizer 1: a block
    // decodes to exactly its sample value.
    private static byte[] WriteFlatBlockJpeg420(int[,] luma, int[,] cb, int[,] cr) {
        int mcuRows = cb.GetLength(0), mcuColumns = cb.GetLength(1);
        byte[] dcBits = { 0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0 };
        var output = new List<byte> { 0xFF, 0xD8, 0xFF, 0xDB, 0, 67, 0 };
        output.AddRange(Enumerable.Repeat((byte)1, 64));
        output.AddRange(new byte[] { 0xFF, 0xC0, 0, 17, 8, 0, (byte)(mcuRows * 16), 0, (byte)(mcuColumns * 16), 3, 1, 0x22, 0, 2, 0x11, 0, 3, 0x11, 0 });
        output.AddRange(new byte[] { 0xFF, 0xC4, 0, 31, 0x00 });
        output.AddRange(dcBits);
        output.AddRange(Enumerable.Range(0, 12).Select(value => (byte)value));
        output.AddRange(new byte[] { 0xFF, 0xC4, 0, 20, 0x10, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0x00 });
        output.AddRange(new byte[] { 0xFF, 0xDA, 0, 12, 3, 1, 0x00, 2, 0x00, 3, 0x00, 0, 63, 0 });

        // Canonical DC codes; the only AC symbol (end of block) is the one-bit code 0.
        var codes = new (int Code, int Length)[12];
        for (int length = 1, code = 0, symbol = 0; length <= 16; length++, code <<= 1) {
            for (int count = 0; count < dcBits[length - 1]; count++) codes[symbol++] = (code++, length);
        }
        long bits = 0L;
        int bitCount = 0;
        void Emit(int value, int length) {
            bits = (bits << length) | (uint)(value & ((1 << length) - 1));
            bitCount += length;
            while (bitCount >= 8) {
                byte next = (byte)(bits >> (bitCount - 8));
                output.Add(next);
                if (next == 0xFF) output.Add(0);
                bitCount -= 8;
            }
        }
        var previous = new int[3];
        void Block(int component, int sample) {
            int dc = (sample - 128) * 8;
            int difference = dc - previous[component];
            previous[component] = dc;
            int category = 0;
            while ((Math.Abs(difference) >> category) != 0) category++;
            Emit(codes[category].Code, codes[category].Length);
            if (category > 0) Emit(difference > 0 ? difference : difference + (1 << category) - 1, category);
            Emit(0, 1);
        }
        for (int row = 0; row < mcuRows; row++) {
            for (int column = 0; column < mcuColumns; column++) {
                Block(0, luma[row * 2, column * 2]);
                Block(0, luma[row * 2, column * 2 + 1]);
                Block(0, luma[row * 2 + 1, column * 2]);
                Block(0, luma[row * 2 + 1, column * 2 + 1]);
                Block(1, cb[row, column]);
                Block(2, cr[row, column]);
            }
        }
        if (bitCount > 0) Emit(0x7F, 8 - bitCount);
        output.AddRange(new byte[] { 0xFF, 0xD9 });
        return output.ToArray();
    }

    private static OfficeRasterImage CreateTiles(int width, int height) {
        var image = new OfficeRasterImage(width, height);
        OfficeColor[] colors = { OfficeColor.FromRgb(33, 81, 171), OfficeColor.FromRgb(181, 35, 93),
            OfficeColor.FromRgb(57, 191, 103), OfficeColor.FromRgb(213, 143, 47) };
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) image.SetPixel(x, y, colors[((y / 16) & 1) * 2 + ((x / 16) & 1)]);
        }
        return image;
    }
}
