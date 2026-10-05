using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingTiffUnsigned16Tests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffUnsigned16");

    public static IEnumerable<object[]> Cases() => File.ReadLines(Path.Combine(Corpus, "manifest.csv"))
        .Skip(1).Select(row => new object[] { row });

    [Theory]
    [MemberData(nameof(Cases))]
    public void IndependentSixteenBitSamplesSurviveCompressionStorageAlphaAndColorConversion(string row) {
        string[] fields = row.Split(',');
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, fields[0]));
        Assert.True(OfficeImageReader.TryValidateContent(bytes, fields[0], out _));
        Assert.True(OfficeRasterContainerInspector.TryInspect(bytes, out var container));
        Assert.Equal(int.Parse(fields[6]), container!.Frames.Count);
        OfficeRasterImage? image;
        if (fields[5] == "1") {
            byte[] profileBytes = OfficeImageMetadataInspector.ReadIccProfile(bytes, OfficeImageFormat.Tiff,
                4 * 1024 * 1024, default, out bool hasProfile)!;
            Assert.True(hasProfile);
            Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out var profile));
            Assert.True(OfficeIccRasterConverter.TryDecodeToSrgb(bytes, profile!, new(), out image));
        } else {
            Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out image));
        }
        Assert.NotNull(image);
        Assert.Equal((int.Parse(fields[1]), int.Parse(fields[2])), (image.Width, image.Height));
        byte[] expected = File.ReadAllBytes(Path.Combine(Corpus, Path.ChangeExtension(fields[0], ".rgba")));
        int tolerance = fields[5] == "1" ? 2 : 0;
        for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
            int offset = (y * image.Width + x) * 4;
            OfficeColor actual = image.GetPixel(x, y);
            Assert.InRange(Math.Abs(actual.R - expected[offset]), 0, tolerance);
            Assert.InRange(Math.Abs(actual.G - expected[offset + 1]), 0, tolerance);
            Assert.InRange(Math.Abs(actual.B - expected[offset + 2]), 0, tolerance);
            Assert.Equal(expected[offset + 3], actual.A);
        }
    }

    [Fact]
    public void SelectedSixteenBitPageAndMultiplePageLossPolicyUseTheSharedInventory() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, "rgb-multipage-be.tif"));
        var options = new OfficeRasterDecodeOptions { FrameIndex = 1 };
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, options, out var image, out var info));
        Assert.Equal(2, info.FrameCount);
        Assert.Equal(1, info.SelectedFrameIndex);
        Assert.Equal((19, 13), (image!.Width, image.Height));
        byte[] expected = File.ReadAllBytes(Path.Combine(Corpus, "rgb-multipage-be.page1.rgba"));
        byte[] firstPage = File.ReadAllBytes(Path.Combine(Corpus, "rgb-multipage-be.rgba"));
        Assert.False(expected.SequenceEqual(firstPage));
        for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
            int offset = (y * image.Width + x) * 4;
            Assert.Equal(OfficeColor.FromRgba(expected[offset], expected[offset + 1],
                expected[offset + 2], expected[offset + 3]), image.GetPixel(x, y));
        }
        options.FrameLossPolicy = OfficeRasterFrameLossPolicy.RejectMultipleFrames;
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, options, out _, out _));
    }

    [Theory]
    [InlineData(258, 8)] // Mixed sample widths are not silently consumed as uniform words.
    [InlineData(266, 2)] // Reversed bit order is outside the managed subset.
    [InlineData(339, 2)]
    [InlineData(339, 3)]
    [InlineData(339, 4)]
    public void UnsupportedSampleDeclarationsRejectValidationAndDecode(int tag, int value) {
        byte[] bytes = PlainRgb();
        SetFirstShort(bytes, tag, value);
        Assert.False(OfficeTiffCodec.TryDecode(bytes, out _));
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "source.tif", out _));
    }

    [Fact]
    public void TruncatedWordPayloadCannotPassValidationOrDecode() {
        byte[] bytes = PlainRgb();
        int entry = Entry(bytes, 279);
        int counts = BitConverter.ToInt32(bytes, entry + 8);
        int count = BitConverter.ToInt32(bytes, counts);
        BitConverter.GetBytes(count - 1).CopyTo(bytes, counts);
        Assert.False(OfficeTiffCodec.TryDecode(bytes, out _));
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "source.tif", out _));
    }

    [Fact]
    public void SixteenBitDecodeAndValidationAccountForBothBytesOfEverySample() {
        byte[] bytes = PlainRgb();
        const int decodedBytes = 19 * 13 * 3 * 2;
        int compressedBytes = decodedBytes; // The fixture is uncompressed.
        Assert.True(OfficeTiffCodec.TryValidateAllPages(bytes, new(), decodedBytes + compressedBytes));
        Assert.False(OfficeTiffCodec.TryValidateAllPages(bytes, new(), decodedBytes + compressedBytes - 1));
        var options = new OfficeRasterDecodeOptions {
            RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes - bytes.Length - decodedBytes - 19 * 13 * 4 - 65536
        };
        Assert.True(OfficeTiffCodec.TryDecodePage(bytes, 0, options, out _));
        options.RetainedManagedBytes++;
        Assert.False(OfficeTiffCodec.TryDecodePage(bytes, 0, options, out _));
        options = new OfficeRasterDecodeOptions { MaximumDecodedPixels = 19 * 13 - 1 };
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, options, out _, out _));
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        options.CancellationToken = cancellation.Token;
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(bytes, options, out _, out _));
    }

    private static byte[] PlainRgb() => File.ReadAllBytes(Path.Combine(Corpus, "rgb-extra-1-none-strip-chunky-le.tif"));

    private static int Entry(byte[] bytes, int tag) {
        int ifd = BitConverter.ToInt32(bytes, 4), count = BitConverter.ToUInt16(bytes, ifd);
        for (int index = 0; index < count; index++) {
            int entry = ifd + 2 + index * 12;
            if (BitConverter.ToUInt16(bytes, entry) == tag) return entry;
        }
        throw new InvalidOperationException("Missing fixture tag.");
    }

    private static void SetFirstShort(byte[] bytes, int tag, int value) {
        int entry = Entry(bytes, tag);
        int target = BitConverter.ToInt32(bytes, entry + 4) * 2 <= 4 ? entry + 8 : BitConverter.ToInt32(bytes, entry + 8);
        bytes[target] = (byte)value; bytes[target + 1] = (byte)(value >> 8);
    }
}
