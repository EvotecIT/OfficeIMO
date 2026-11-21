using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Managed JPEG 2000 decoding against independently decoded samples (OpenJPEG) of files written by JasPer,
/// Kakadu and OpenJPEG. See OfficeIMO.TestAssets/Images/Jpeg2000/PROVENANCE.md.
/// </summary>
public sealed class Jpeg2000DecoderTests {
    private static readonly string Folder = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Images", "Jpeg2000");

    public static TheoryData<string> Fixtures() {
        var data = new TheoryData<string>();
        foreach (JsonElement entry in ReadManifest().GetProperty("files").EnumerateArray()) data.Add(entry.GetProperty("file").GetString()!);
        return data;
    }

    [Theory]
    [MemberData(nameof(Fixtures))]
    public void DecodesFixtureLikeTheReferenceDecoder(string file) {
        JsonElement entry = ReadManifest().GetProperty("files").EnumerateArray().Single(item => item.GetProperty("file").GetString() == file);
        byte[] bytes = File.ReadAllBytes(Path.Combine(Folder, file));
        Assert.Equal(entry.GetProperty("sha256").GetString(), Sha256(bytes));

        Assert.True(OfficeJpeg2000Decoder.TryDecode(bytes, new OfficeJpeg2000DecodeOptions(), out OfficeJpeg2000Image? image, out string? failure), failure);
        if (entry.TryGetProperty("samplesSha256", out JsonElement expected)) {
            Assert.Equal(expected.GetString(), Sha256(image!.Samples));
            return;
        }
        // Irreversible streams: compare with the lossless decode of the same source image.
        byte[] lossless = File.ReadAllBytes(Path.Combine(Folder, entry.GetProperty("lossyOf").GetString()!));
        Assert.True(OfficeJpeg2000Decoder.TryDecode(lossless, new OfficeJpeg2000DecodeOptions(), out OfficeJpeg2000Image? source));
        Assert.Equal(source!.Samples.Length, image!.Samples.Length);
        int maximum = entry.GetProperty("maximumDifference").GetInt32();
        for (int index = 0; index < image.Samples.Length; index++) Assert.InRange(Math.Abs(image.Samples[index] - source.Samples[index]), 0, maximum);
    }

    [Fact]
    public void RasterDecoderUsesManagedSamplesAndPalettes() {
        byte[] indexed = File.ReadAllBytes(Path.Combine(Folder, "openjpeg-lossless-indexed-u8-rgb-u8.jp2"));
        // The raster pipeline admits JP2 files whose headers pass the bounded opaque-subset inspection.
        byte[] rgb = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "red-rgb.jp2"));
        Assert.True(OfficeRasterImageDecoder.TryDecode(rgb, new OfficeRasterDecodeOptions(), out OfficeRasterImage? raster, out _));
        Assert.True(OfficeJpeg2000Decoder.TryDecode(rgb, new OfficeJpeg2000DecodeOptions(), out OfficeJpeg2000Image? samples));
        Assert.Equal((samples!.Width, samples.Height), (raster!.Width, raster.Height));
        for (int pixel = 0; pixel < samples.Width * samples.Height; pixel++) {
            for (int channel = 0; channel < 3; channel++) Assert.Equal(samples.Samples[pixel * 3 + channel], raster.PixelBuffer[pixel * 4 + channel]);
            Assert.Equal(255, raster.PixelBuffer[pixel * 4 + 3]);
        }

        Assert.True(OfficeJpeg2000Decoder.TryDecode(indexed, new OfficeJpeg2000DecodeOptions(), out OfficeJpeg2000Image? palette));
        Assert.True(palette!.UsedPalette);
        Assert.Equal(3, palette.Channels);
        // Without the container channels the palette indices themselves are returned (PDF /Indexed images).
        Assert.True(OfficeJpeg2000Decoder.TryDecode(indexed,
            new OfficeJpeg2000DecodeOptions { ApplyContainerChannels = false, ScaleToEightBits = false }, out OfficeJpeg2000Image? indices));
        Assert.Equal(1, indices!.Channels);
    }

    [Fact]
    public void RejectsImagesBeyondPixelAndMemoryBudgets() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Folder, "jasper-tile4x2-res5.jp2"));
        Assert.False(OfficeJpeg2000Decoder.TryDecode(bytes, new OfficeJpeg2000DecodeOptions { MaximumPixels = 1000 }, out _));
        Assert.False(OfficeJpeg2000Decoder.TryDecode(bytes, new OfficeJpeg2000DecodeOptions { MaximumWorkingBytes = 64 * 1024 }, out _));
        Assert.ThrowsAny<OperationCanceledException>(() => OfficeJpeg2000Decoder.TryDecode(bytes,
            new OfficeJpeg2000DecodeOptions { CancellationToken = new CancellationToken(true) }, out _));
    }

    // Untrusted input: truncations and byte corruptions must fail closed or decode, never throw or run away.
    [Theory]
    [InlineData("jasper-rgba-u8-cbstyle-63-all.jp2")]
    [InlineData("openjpeg-lossless-rgba-u8-prog3-tile4x2-cblk4x16-tp3-layers3-res2.jp2")]
    [InlineData("kakadu-lossy-rgba-u8-prog0-layers1-res6-mct.jp2")]
    public void MalformedInputFailsClosed(string file) {
        byte[] original = File.ReadAllBytes(Path.Combine(Folder, file));
        var random = new Random(file.Length);
        var options = new OfficeJpeg2000DecodeOptions { MaximumWorkingBytes = 32L * 1024 * 1024 };
        for (int length = 0; length < original.Length; length += Math.Max(1, original.Length / 97))
            OfficeJpeg2000Decoder.TryDecode(original.Take(length).ToArray(), options, out _);
        for (int round = 0; round < 300; round++) {
            byte[] mutated = (byte[])original.Clone();
            int changes = 1 + random.Next(8);
            for (int change = 0; change < changes; change++) mutated[random.Next(Math.Min(mutated.Length, 400))] = (byte)random.Next(256);
            if (random.Next(2) == 0) mutated[random.Next(mutated.Length)] ^= (byte)(1 << random.Next(8));
            OfficeJpeg2000Decoder.TryDecode(mutated, options, out _);
        }
    }

    private static string Sha256(byte[] data) {
        using var hash = SHA256.Create();
        return BitConverter.ToString(hash.ComputeHash(data)).Replace("-", string.Empty).ToLowerInvariant();
    }

    private static JsonElement ReadManifest() =>
        JsonDocument.Parse(File.ReadAllText(Path.Combine(Folder, "manifest.json"))).RootElement;
}
