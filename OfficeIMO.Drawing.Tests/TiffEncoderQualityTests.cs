using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class TiffEncoderQualityTests {
    private const string Eofb = "000000000001000000000001";

    [Fact]
    public void Group4WhiteRowsUseVerticalCodesAndEndOfFacsimileBlock() {
        var image = new OfficeRasterImage(19, 3, OfficeColor.White);
        byte[] encoded = OfficeTiffCodec.Encode(image, Group4());
        Assert.Equal(Bits("111" + Eofb), Strip(encoded));
        Assert.Equal(1, Scalar(encoded, 258));
        Assert.Equal(4, Scalar(encoded, 259));
        Assert.Equal(0, Scalar(encoded, 262));
        Assert.Equal(1, Scalar(encoded, 277));
        Assert.Equal(0, Scalar(encoded, 293));
        Assert.Equal(-1, Entry(encoded, 317));
        Assert.Equal(-1, Entry(encoded, 338));
        AssertDecoded(image, encoded);
    }

    [Theory]
    [InlineData(-3, "0000010")]
    [InlineData(-2, "000010")]
    [InlineData(-1, "010")]
    [InlineData(0, "1")]
    [InlineData(1, "011")]
    [InlineData(2, "000011")]
    [InlineData(3, "0000011")]
    public void Group4UsesNormativeHorizontalAndSignedVerticalCodes(int shift, string vertical) {
        var image = new OfficeRasterImage(24, 2, OfficeColor.White);
        PaintRun(image, 0, 8, 16);
        PaintRun(image, 1, 8 + shift, 16 + shift);
        // First row: H, white 8 (10011), black 8 (000101), V(0).
        // Second row: two shifted edges, then V(0) at the row boundary.
        const string first = "001100110001011";
        byte[] encoded = OfficeTiffCodec.Encode(image, Group4());
        Assert.Equal(Bits(first + vertical + vertical + "1" + Eofb), Strip(encoded));
        AssertDecoded(image, encoded);
    }

    [Fact]
    public void Group4UsesPassModeWhenReferenceBlackRunDisappears() {
        var image = new OfficeRasterImage(24, 2, OfficeColor.White);
        PaintRun(image, 0, 8, 16);
        byte[] encoded = OfficeTiffCodec.Encode(image, Group4());
        Assert.Equal(Bits("001100110001011" + "00011" + Eofb), Strip(encoded));
        AssertDecoded(image, encoded);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(7)]
    [InlineData(8)]
    [InlineData(9)]
    [InlineData(83)]
    [InlineData(257)]
    [InlineData(5201)]
    public void Group4RetainsPolarityUnalignedRowsAndLongMakeupRuns(int width) {
        var image = Pattern(width, 19);
        byte[] encoded = OfficeTiffCodec.Encode(image, Group4());
        AssertDecoded(image, encoded);
        Assert.True(OfficeImageReader.TryValidateContent(encoded, "bilevel.tif", out _));
        using var stream = new NonSeekableStream();
        OfficeTiffCodec.EncodeTo(image, stream, Group4());
        Assert.True(stream.CanWrite);
        Assert.Equal(encoded, stream.ToArray());
        Assert.Equal(encoded, OfficeTiffCodec.EncodePages(new[] { image }, Group4()));
    }

    [Fact]
    public void Group4RetainsEveryFivePixelRowAndItsReferenceTransitions() {
        // All 32 row patterns cover both initial colors, final colors, and zero-width initial white runs.
        var image = new OfficeRasterImage(5, 32, OfficeColor.White);
        for (int y = 0; y < 32; y++) for (int x = 0; x < 5; x++)
            if ((y & (1 << x)) != 0) image.SetPixel(x, y, OfficeColor.Black);
        AssertDecoded(image, OfficeTiffCodec.Encode(image, Group4()));
    }

    [Theory]
    [InlineData(OfficeTiffCompression.None)]
    [InlineData(OfficeTiffCompression.PackBits)]
    [InlineData(OfficeTiffCompression.Deflate)]
    [InlineData(OfficeTiffCompression.Lzw)]
    [InlineData(OfficeTiffCompression.CcittGroup4)]
    public void MultipageWritersRetainPageDimensionsPixelsAndNativeResolution(OfficeTiffCompression compression) {
        var pages = new[] { Pattern(13, 9), Pattern(29, 7) };
        var options = new OfficeTiffEncodeOptions {
            Compression = compression,
            Resolution = new OfficeImageResolution(118.25, 59.125, OfficeImageResolutionUnit.PixelsPerCentimeter)
        };
        byte[] encoded = OfficeTiffCodec.EncodePages(pages, options);
        Assert.True(OfficeTiffCodec.TryGetPageCount(encoded, out int pageCount));
        Assert.Equal(2, pageCount);
        for (int page = 0; page < pages.Length; page++) {
            AssertDecoded(pages[page], encoded, page);
            Assert.Equal((int)compression, Scalar(encoded, 259, page));
            Assert.Equal(3, Scalar(encoded, 296, page));
            Assert.Equal(118.25, Rational(encoded, 282, page));
            Assert.Equal(59.125, Rational(encoded, 283, page));
        }
        using var stream = new NonSeekableStream();
        OfficeTiffCodec.EncodePagesTo(pages, stream, options);
        Assert.Equal(encoded, stream.ToArray());
    }

    [Theory]
    [InlineData(OfficeImageResolutionUnit.AspectRatio, 3.0, 2.0, 1, 3.0, 2.0)]
    [InlineData(OfficeImageResolutionUnit.PixelsPerInch, 300.5, 150.25, 2, 300.5, 150.25)]
    [InlineData(OfficeImageResolutionUnit.PixelsPerMeter, 12000.0, 6000.0, 3, 120.0, 60.0)]
    public void Group4SharesNativeDensityAcrossByteStreamAndPageWriters(OfficeImageResolutionUnit unit,
        double x, double y, int tiffUnit, double nativeX, double nativeY) {
        var image = Pattern(17, 7);
        var options = Group4();
        options.Resolution = new OfficeImageResolution(x, y, unit);
        using var stream = new MemoryStream();
        OfficeTiffCodec.EncodeTo(image, stream, options);
        foreach (byte[] encoded in new[] { OfficeTiffCodec.Encode(image, options), stream.ToArray(), OfficeTiffCodec.EncodePages(new[] { image }, options) }) {
            Assert.Equal(tiffUnit, Scalar(encoded, 296));
            Assert.Equal(nativeX, Rational(encoded, 282));
            Assert.Equal(nativeY, Rational(encoded, 283));
            AssertDecoded(image, encoded);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Group4UsesCentralIccCarrierAndDensitySuppression(bool writeResolution) {
        var image = Pattern(17, 7);
        byte[] icc = new byte[132]; // Structurally valid empty ICC profile; no color transform is requested.
        icc[3] = 132;
        Encoding.ASCII.GetBytes("acsp", 0, 4, icc, 36);
        var metadata = new OfficeImageMetadata {
            IccProfile = icc,
            Resolution = new OfficeImageResolution(200, 100)
        };
        var options = new OfficeRasterEncodingOptions { Tiff = Group4(), WriteResolutionMetadata = writeResolution };
        byte[] encoded = OfficeRasterImageEncoder.EncodeWithMetadata(image, OfficeImageExportFormat.Tiff, metadata, options).RequireMetadataPreservation();
        Assert.Equal(icc, OfficeImageMetadata.Read(encoded).IccProfile);
        Assert.Equal(4, Scalar(encoded, 259));
        AssertDecoded(image, encoded);
        if (writeResolution) {
            Assert.Equal(200, Rational(encoded, 282));
            Assert.Equal(100, Rational(encoded, 283));
            Assert.Equal(2, Scalar(encoded, 296));
        } else {
            Assert.Equal(-1, Entry(encoded, 282));
            Assert.Equal(-1, Entry(encoded, 283));
            Assert.Equal(-1, Entry(encoded, 296));
        }
    }

    [Fact]
    public void Group4CanOmitNativeDensityOnEveryPage() {
        var image = Pattern(19, 7);
        var options = Group4();
        options.WriteResolution = false;
        using var stream = new MemoryStream();
        OfficeTiffCodec.EncodeTo(image, stream, options);
        foreach (byte[] encoded in new[] { OfficeTiffCodec.Encode(image, options), stream.ToArray(), OfficeTiffCodec.EncodePages(new[] { image, image }, options) }) {
            Assert.True(OfficeTiffCodec.TryGetPageCount(encoded, out int pages));
            for (int page = 0; page < pages; page++) {
                AssertDecoded(image, encoded, page);
                Assert.Equal(-1, Entry(encoded, 282, page));
                Assert.Equal(-1, Entry(encoded, 283, page));
                Assert.Equal(-1, Entry(encoded, 296, page));
            }
        }
    }

    [Theory]
    [InlineData(128, 128, 128, 255)]
    [InlineData(0, 255, 0, 255)]
    [InlineData(0, 0, 0, 0)]
    [InlineData(255, 255, 255, 254)]
    public void Group4RejectsNonBilevelOrAlphaSamplesBeforeDestinationWrites(byte r, byte g, byte b, byte a) {
        var image = Pattern(19, 7);
        image.SetPixel(18, 6, OfficeColor.FromRgba(r, g, b, a));
        Assert.Throws<ArgumentException>(() => OfficeTiffCodec.Encode(image, Group4()));
        using var stream = new MemoryStream();
        Assert.Throws<ArgumentException>(() => OfficeTiffCodec.EncodeTo(image, stream, Group4()));
        Assert.Equal(0, stream.Length);
        Assert.Throws<ArgumentException>(() => OfficeTiffCodec.EncodePagesTo(new[] { Pattern(7, 5), image }, stream, Group4()));
        Assert.Equal(0, stream.Length);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(9)]
    public void Group4CancellationInCountingOrWritingPassLeavesDestinationUntouched(int cancelAtRow) {
        using var cancellation = new CancellationTokenSource();
        using var stream = new MemoryStream();
        int rows = 0;
        Assert.Throws<OperationCanceledException>(() => OfficeTiffCodec.EncodeTo(Pattern(83, 7), stream,
            Group4(), cancellation.Token, checkpoint => {
                if (checkpoint == OfficeRasterEncodingCheckpoint.TiffCompressionRow && ++rows == cancelAtRow)
                    cancellation.Cancel();
            }));
        Assert.Equal(0, stream.Length);
        Assert.True(stream.CanWrite);
    }

    private static OfficeTiffEncodeOptions Group4() => new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.CcittGroup4 };

    private static OfficeRasterImage Pattern(int width, int height) {
        var image = new OfficeRasterImage(width, height, OfficeColor.White);
        for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
            bool black = y % 7 == 1 || y % 7 == 2 && (x & 1) == 0 ||
                y % 7 == 3 && (x / 13 & 1) != 0 || y % 7 == 4 && x > width / 3 && x < width * 2 / 3 ||
                y % 7 == 5 && ((x * 1103515245L + y * 12345L) >> 15 & 1) != 0;
            if (black) image.SetPixel(x, y, OfficeColor.Black);
        }
        return image;
    }

    private static void PaintRun(OfficeRasterImage image, int row, int start, int end) {
        for (int x = start; x < end; x++) image.SetPixel(x, row, OfficeColor.Black);
    }

    private static void AssertDecoded(OfficeRasterImage expected, byte[] bytes, int page = 0) {
        Assert.True(OfficeTiffCodec.TryDecodePage(bytes, page, out OfficeRasterImage? actual));
        Assert.Equal(expected.Width, actual!.Width);
        Assert.Equal(expected.Height, actual.Height);
        Assert.Equal(expected.GetPixels(), actual.GetPixels());
    }

    private static int Entry(byte[] bytes, int tag, int page = 0) {
        int ifd = BitConverter.ToInt32(bytes, 4);
        while (page-- > 0) ifd = BitConverter.ToInt32(bytes, ifd + 2 + BitConverter.ToUInt16(bytes, ifd) * 12);
        for (int i = 0; i < BitConverter.ToUInt16(bytes, ifd); i++) {
            int entry = ifd + 2 + i * 12;
            if (BitConverter.ToUInt16(bytes, entry) == tag) return entry;
        }
        return -1;
    }

    private static int Scalar(byte[] bytes, int tag, int page = 0) {
        int entry = Entry(bytes, tag, page);
        Assert.True(entry >= 0);
        return BitConverter.ToUInt16(bytes, entry + 2) == 3
            ? BitConverter.ToUInt16(bytes, entry + 8) : BitConverter.ToInt32(bytes, entry + 8);
    }

    private static double Rational(byte[] bytes, int tag, int page = 0) {
        int entry = Entry(bytes, tag, page);
        Assert.True(entry >= 0);
        int offset = BitConverter.ToInt32(bytes, entry + 8);
        return BitConverter.ToUInt32(bytes, offset) / (double)BitConverter.ToUInt32(bytes, offset + 4);
    }

    private static byte[] Strip(byte[] bytes) => bytes.Skip(Scalar(bytes, 273)).Take(Scalar(bytes, 279)).ToArray();

    private static byte[] Bits(string bits) {
        bits = bits.PadRight((bits.Length + 7) / 8 * 8, '0');
        return Enumerable.Range(0, bits.Length / 8).Select(i => Convert.ToByte(bits.Substring(i * 8, 8), 2)).ToArray();
    }

    private sealed class NonSeekableStream : MemoryStream {
        public override bool CanSeek => false;
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    }
}
