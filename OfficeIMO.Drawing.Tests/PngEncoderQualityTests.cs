using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PngEncoderQualityTests {
    [Theory]
    [InlineData(0, 0, 1)]
    [InlineData(1, 0, 2)]
    [InlineData(2, 0, 4)]
    [InlineData(3, 0, 8)]
    [InlineData(4, 4, 8)]
    [InlineData(5, 3, 1)]
    [InlineData(6, 2, 8)]
    [InlineData(7, 6, 8)]
    public void ExactSampleLayoutsRetainEveryChannelAndAreDeterministic(int fixture, int type, int depth) {
        var image = new OfficeRasterImage(513, 5);
        for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
            byte g = fixture == 0 ? (byte)(x % 2 * 255) : fixture == 1 ? (byte)(x % 4 * 85) : fixture == 2 ? (byte)(x % 16 * 17) : (byte)x;
            byte r = g, b = g, a = 255;
            if (fixture == 4) a = (byte)y;
            if (fixture == 5) { r = (byte)(x % 2 * 140); g = 71; b = 233; a = (byte)(x % 2 * 255); }
            if (fixture >= 6) { g = (byte)y; b = 19; if (fixture == 7) a = (byte)(x + y); }
            image.SetPixel(x, y, OfficeColor.FromRgba(r, g, b, a));
        }
        byte[] png = OfficePngWriter.Encode(image);
        Assert.Equal(type, png[25]); Assert.Equal(depth, png[24]);
        Assert.Equal(png, OfficePngWriter.Encode(image));
        using var output = new MemoryStream(); OfficePngWriter.EncodeTo(image, output);
        Assert.Equal(png, output.ToArray());
        Assert.True(OfficePngReader.TryDecode(png, out var decoded));
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
        Assert.Contains("sRGB", ChunkTypes(png)); Assert.Contains("gAMA", ChunkTypes(png)); Assert.Contains("cHRM", ChunkTypes(png));
        if (fixture == 5) Assert.Contains("tRNS", ChunkTypes(png));
    }

    [Theory]
    [InlineData(false, false, 6)]
    [InlineData(false, true, 2)]
    [InlineData(true, false, 4)]
    [InlineData(true, true, 0)]
    public void SixteenBitEntryPointRetainsLowBitsInNetworkByteOrder(bool gray, bool opaque, int type) {
        ushort[] samples = { 0x0102, gray ? (ushort)0x0102 : (ushort)0x0304, gray ? (ushort)0x0102 : (ushort)0x0506, opaque ? ushort.MaxValue : (ushort)0x0708 };
        byte[] png = OfficePngWriter.EncodeRgba64(1, 1, samples);
        Assert.Equal(16, png[24]); Assert.Equal(type, png[25]);
        using var idat = new MemoryStream();
        for (int p = 8; p < png.Length;) {
            int n = ReadInt(png, p); if (Encoding.ASCII.GetString(png, p + 4, 4) == "IDAT") idat.Write(png, p + 8, n); p += n + 12;
        }
        // Every PNG filter predicts zero for this first-row, single-pixel sample.
        idat.Position = 2; using var inflate = new System.IO.Compression.DeflateStream(idat, System.IO.Compression.CompressionMode.Decompress);
        using var raw = new MemoryStream(); inflate.CopyTo(raw); byte[] row = raw.ToArray();
        var expected = gray ? opaque ? new byte[] { 1, 2 } : new byte[] { 1, 2, 7, 8 } : opaque ? new byte[] { 1, 2, 3, 4, 5, 6 } : new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 };
        Assert.Equal(expected, row.Skip(1));
    }

    [Fact]
    public void RgbProfileKeepsGrayPixelsInAnRgbCompatibleLayout() {
        byte[] profile = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "IccColorCorpus", "littlecms-rgb-matrix.icc"));
        var image = new OfficeRasterImage(257, 2);
        for (int x = 0; x < image.Width; x++) for (int y = 0; y < image.Height; y++)
            image.SetPixel(x, y, OfficeColor.FromRgb((byte)x, (byte)x, (byte)x));
        var metadata = new OfficeImageMetadata { IccProfile = profile };
        byte[] png = OfficeRasterImageEncoder.EncodeWithMetadata(image, OfficeImageExportFormat.Png, metadata).EncodedBytes;
        Assert.Contains((int)png[25], new[] { 2, 3, 6 });
        Assert.Contains("iCCP", ChunkTypes(png));
        Assert.DoesNotContain("sRGB", ChunkTypes(png));
        Assert.DoesNotContain("gAMA", ChunkTypes(png));
        Assert.DoesNotContain("cHRM", ChunkTypes(png));
        Assert.True(OfficePngReader.TryDecode(png, out var decoded));
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
    }

    [Theory]
    [InlineData(OfficePngCompression.Optimal)]
    [InlineData(OfficePngCompression.Fastest)]
    [InlineData(OfficePngCompression.Stored)]
    public void GrayProfileUsesGraySamplesInsteadOfAnIncompatiblePalette(OfficePngCompression compression) {
        var metadata = OfficeImageMetadata.Read(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,
            "TestAssets", "TiffUnsigned16", "gray-icc-alpha-le.tif")));
        Assert.NotNull(metadata.IccProfile);
        var image = new OfficeRasterImage(64, 8, OfficeColor.FromRgba(17, 17, 17, 101));
        var options = new OfficeRasterEncodingOptions { Png = new OfficePngEncodeOptions { Compression = compression } };
        byte[] png = OfficeRasterImageEncoder.EncodeWithMetadata(image, OfficeImageExportFormat.Png, metadata, options).EncodedBytes;
        Assert.Equal(4, png[25]);
        Assert.True(OfficePngReader.TryDecode(png, out var decoded));
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
        image.SetPixel(0, 0, OfficeColor.Red);
        Assert.Throws<ArgumentException>(() => OfficeRasterImageEncoder.EncodeWithMetadata(image, OfficeImageExportFormat.Png, metadata, options));
    }

    private static string[] ChunkTypes(byte[] png) {
        var names = new System.Collections.Generic.List<string>();
        for (int p = 8; p < png.Length;) { names.Add(Encoding.ASCII.GetString(png, p + 4, 4)); p += ReadInt(png, p) + 12; }
        return names.ToArray();
    }
    private static int ReadInt(byte[] bytes, int p) => bytes[p] << 24 | bytes[p + 1] << 16 | bytes[p + 2] << 8 | bytes[p + 3];
}
