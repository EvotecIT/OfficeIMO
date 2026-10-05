using OfficeIMO.Drawing;
using Xunit;
using System.Threading;

namespace OfficeIMO.Tests;

public sealed class DrawingTiffFloatingTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffFloating");

    [Fact]
    public void IndependentFloatingSamplesSurviveCompressionByteOrderStorageAndAlpha() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string name = row.Split(',')[0];
            byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, name));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, name, out _), name);
            Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out var image), name);
            Assert.Equal((19, 17), (image!.Width, image.Height));
            Assert.Equal(File.ReadAllBytes(Path.Combine(Corpus, Path.ChangeExtension(name, ".rgba"))), image.GetPixels());
        }
    }

    [Theory]
    [InlineData(16, 0x7C00UL)]
    [InlineData(16, 0x7E00UL)]
    [InlineData(32, 0x7F800000UL)]
    [InlineData(32, 0x7FC00000UL)]
    [InlineData(64, 0x7FF0000000000000UL)]
    [InlineData(64, 0x7FF8000000000000UL)]
    public void NonfiniteColorOrAlphaFailsWithoutPartialPixels(int bits, ulong value) {
        foreach (bool planar in new[] { false, true }) foreach (bool tiled in new[] { false, true })
        foreach (int channel in new[] { 0, 3 }) {
            byte[] bytes = StorageFixture(bits, planar, tiled);
            int offset = SampleOffset(bytes, bits, planar, tiled, channel);
            for (int i = 0; i < bits / 8; i++) bytes[offset + i] = (byte)(value >> (i * 8));
            Assert.False(OfficeImageReader.TryValidateContent(bytes, "nonfinite.tif", out _));
            Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, out var image));
            Assert.Null(image);
        }
    }

    [Fact]
    public void SourcePrecisionSurvivesLowAssociatedAlphaAndSdrClipping() {
        byte[] bytes = Fixture(64, 1);
        int offset = FirstStrip(bytes);
        double alpha = 1D / 65536;
        foreach (var pair in new[] { (0, .5 * alpha), (1, 1.5 * alpha), (2, -.5 * alpha), (3, alpha) })
            BitConverter.GetBytes(pair.Item2).CopyTo(bytes, offset + pair.Item1 * 8);
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out var image));
        Assert.Equal(OfficeColor.FromRgba(128, 255, 0, 0), image!.GetPixel(0, 0));
        byte[] profileBytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "IccColorCorpus", "littlecms-rgb-matrix.icc"));
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out var profile));
        Assert.True(profile!.TryConvert(new[] { .5, 1.5, -.5 }, OfficeIccRenderingIntent.RelativeColorimetric, out var expected));
        Assert.True(OfficeIccRasterConverter.TryDecodeToSrgb(bytes, profile, new(), out var converted));
        Assert.Equal(OfficeColor.FromRgba(expected.R, expected.G, expected.B, 0), converted!.GetPixel(0, 0));
        foreach (var pair in new[] { (0, double.MaxValue), (1, 0D), (2, 0D), (3, double.Epsilon) })
            BitConverter.GetBytes(pair.Item2).CopyTo(bytes, offset + pair.Item1 * 8);
        Assert.True(profile.TryConvert(new[] { 1D, 0D, 0D }, OfficeIccRenderingIntent.RelativeColorimetric, out expected));
        Assert.True(OfficeIccRasterConverter.TryDecodeToSrgb(bytes, profile, new(), out converted));
        Assert.Equal(OfficeColor.FromRgba(expected.R, expected.G, expected.B, 0), converted!.GetPixel(0, 0));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 322 }, out _, out _));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token }, out _, out _));
    }

    [Theory]
    [InlineData(16, 0x7C00UL)]
    [InlineData(32, 0x7FC00000UL)]
    [InlineData(64, 0x7FF0000000000000UL)]
    public void UnspecifiedExtrasAndTilePaddingAreIgnoredEvenWhenNonfinite(int bits, ulong value) {
        foreach (bool planar in new[] { false, true }) foreach (bool tiled in new[] { false, true }) {
            byte[] bytes = StorageFixture(bits, planar, tiled);
            int extra = Entry(bytes, 338);
            bytes[extra + 8] = 0; bytes[extra + 9] = 0;
            int offset = SampleOffset(bytes, bits, planar, tiled, 3);
            for (int i = 0; i < bits / 8; i++) bytes[offset + i] = (byte)(value >> (i * 8));
            if (tiled) {
                int entry = Entry(bytes, 324), offsets = BitConverter.ToInt32(bytes, entry + 8);
                // Second tile has three visible columns; its last column is padding.
                int padding = BitConverter.ToInt32(bytes, offsets + 4) + 15 * (planar ? 1 : 4) * bits / 8;
                for (int i = 0; i < bits / 8; i++) bytes[padding + i] = (byte)(value >> (i * 8));
                int bottomPadding = BitConverter.ToInt32(bytes, offsets + 12) + 15 * 16 * (planar ? 1 : 4) * bits / 8;
                for (int i = 0; i < bits / 8; i++) bytes[bottomPadding + i] = (byte)(value >> (i * 8));
            }
            Assert.True(OfficeImageReader.TryValidateContent(bytes, "ignored.tif", out _));
            Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out var image));
            byte[] expected = File.ReadAllBytes(Path.Combine(Corpus,
                $"f{bits}-le-planar{(planar ? 1 : 0)}-tile{(tiled ? 1 : 0)}-c1-a2.rgba"));
            for (int i = 3; i < expected.Length; i += 4) expected[i] = 255;
            Assert.Equal(expected, image!.GetPixels());
        }
    }

    [Theory]
    [InlineData(317, 2)] // Integer differencing must not be applied to floating samples.
    [InlineData(339, 1)] // Floating prediction must not be applied to unsigned words.
    public void PredictorMustMatchTheSampleRepresentation(int tag, int value) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, "f16-le-planar0-tile0-c8-a2.tif"));
        int ifd = BitConverter.ToInt32(bytes, 4), count = BitConverter.ToUInt16(bytes, ifd);
        bool changed = false;
        for (int i = 0; i < count; i++) {
            int entry = ifd + 2 + 12 * i;
            if (BitConverter.ToUInt16(bytes, entry) != tag) continue;
            int samples = BitConverter.ToInt32(bytes, entry + 4);
            int offset = samples <= 2 ? entry + 8 : BitConverter.ToInt32(bytes, entry + 8);
            for (int c = 0; c < samples; c++) { bytes[offset + c * 2] = (byte)value; bytes[offset + c * 2 + 1] = 0; }
            changed = true;
        }
        Assert.True(changed);
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "invalid.tif", out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, out _));
    }

    private static byte[] Fixture(int bits, int alpha = 2) =>
        File.ReadAllBytes(Path.Combine(Corpus, $"f{bits}-le-planar0-tile0-c1-a{alpha}.tif"));

    private static byte[] StorageFixture(int bits, bool planar, bool tiled) =>
        File.ReadAllBytes(Path.Combine(Corpus, $"f{bits}-le-planar{(planar ? 1 : 0)}-tile{(tiled ? 1 : 0)}-c1-a2.tif"));

    private static int SampleOffset(byte[] bytes, int bits, bool planar, bool tiled, int channel) {
        int entry = Entry(bytes, tiled ? 324 : 273);
        int count = BitConverter.ToInt32(bytes, entry + 4), offsets = BitConverter.ToInt32(bytes, entry + 8);
        int segment = planar ? channel * (count / 4) : 0;
        return BitConverter.ToInt32(bytes, offsets + segment * 4) + (planar ? 0 : channel * bits / 8);
    }

    private static int Entry(byte[] bytes, int tag) {
        int ifd = BitConverter.ToInt32(bytes, 4), count = BitConverter.ToUInt16(bytes, ifd);
        for (int i = 0; i < count; i++) {
            int entry = ifd + 2 + 12 * i;
            if (BitConverter.ToUInt16(bytes, entry) == tag) return entry;
        }
        throw new InvalidOperationException("Fixture tag is missing.");
    }

    private static int FirstStrip(byte[] bytes) {
        int ifd = BitConverter.ToInt32(bytes, 4), count = BitConverter.ToUInt16(bytes, ifd);
        for (int i = 0; i < count; i++) {
            int entry = ifd + 2 + 12 * i;
            if (BitConverter.ToUInt16(bytes, entry) == 273)
                return BitConverter.ToInt32(bytes, BitConverter.ToInt32(bytes, entry + 8));
        }
        throw new InvalidOperationException("Strip fixture missing offsets.");
    }
}
