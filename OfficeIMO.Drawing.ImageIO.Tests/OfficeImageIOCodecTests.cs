using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Drawing.ImageIO.Tests;

[CollectionDefinition("ImageIO registration", DisableParallelization = true)]
public sealed class ImageIORegistrationCollection { }

[Collection("ImageIO registration")]
public sealed class OfficeImageIOCodecTests {
    private readonly ITestOutputHelper _output;
    public OfficeImageIOCodecTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public void RegistrationAndAvailabilityFollowTheOperatingSystem() {
        IOfficeRasterImageCodec? previous = OfficeRasterImageCodecs.Platform;
        var sentinel = new OfficeImageIOCodec();
        try {
            OfficeRasterImageCodecs.Platform = sentinel;
            OfficeImageIOCodec.Register();
            var availability = sentinel.GetAvailability();
            Assert.Equal(OfficeImageIOCodec.IsSupported, availability.IsAvailable);
            if (OfficeImageIOCodec.IsSupported) {
                Assert.IsType<OfficeImageIOCodec>(OfficeRasterImageCodecs.Platform);
                Assert.NotSame(sentinel, OfficeRasterImageCodecs.Platform);
            } else {
                Assert.Same(sentinel, OfficeRasterImageCodecs.Platform);
                Assert.Equal(OfficeRasterCodecUnavailableReason.NoPlatformCodec, availability.Reason);
                Assert.False(sentinel.TryDecode(new byte[] { 1, 2, 3 }, "image/heic", out var image));
                Assert.Null(image);
            }
        } finally { OfficeRasterImageCodecs.Platform = previous; }
    }

    [Fact]
    public void InvalidInputAndInvalidBudgetsCannotProduceAnImage() {
        var codec = new OfficeImageIOCodec(maximumEncodedBytes: 16);
        foreach (byte[]? bytes in new byte[]?[] { null, Array.Empty<byte>(), new byte[17], new byte[] { 1, 2, 3 } }) {
            Assert.False(codec.TryDecode(bytes!, null, out var image));
            Assert.Null(image);
        }
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeImageIOCodec(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeImageIOCodec(50_000_001));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeImageIOCodec(maximumEncodedBytes: 0));
    }

    [SkippableTheory]
    [InlineData(false)]
    [InlineData(true)]
    public void KnownPngSamplesRetainTheirSourceEncodingAndRowOrder(bool displayP3) {
        RequireMac();
        byte[] expected = AppleImageIOReference.ColorfulRgba();
        byte[] png = AppleImageIOReference.Encode(expected, 32, 24, "public.png", displayP3);
        _output.WriteLine(AppleImageIOReference.DescribeStoredImage(png));
        Assert.True(new OfficeImageIOCodec().TryDecode(png, "image/png", out var raw));
        Assert.NotNull(raw);
        Assert.Equal(expected, raw.GetPixels());
        if (displayP3) {
            byte[] matched = AppleImageIOReference.DecodeFullFile(png).GetPixels();
            Assert.True(expected.Zip(matched, (a, b) => Math.Abs(a - b)).Max() > 10,
                "The wide-gamut control must distinguish preserved samples from sRGB color matching.");
        }
    }

    [SkippableFact]
    public void NativePremultiplicationIsReturnedAsStraightAlpha() {
        RequireMac();
        byte[] expected = { 160, 80, 40, 128, 20, 60, 100, 255, 80, 90, 100, 0 };
        byte[] png = AppleImageIOReference.Encode(expected, 3, 1, "public.png");
        Assert.True(new OfficeImageIOCodec().TryDecode(png, null, out var image));
        byte[] actual = image!.GetPixels();
        for (int i = 0; i < 8; i++) Assert.InRange(Math.Abs(expected[i] - actual[i]), 0, 1);
        Assert.Equal(new byte[4], actual.Skip(8).ToArray());
    }

    [SkippableFact]
    public void MonochromeSourceSamplesExpandToRgbWithoutApplyingTheGrayProfile() {
        RequireMac();
        byte[] pairs = { 32, 255, 160, 128, 224, 255, 96, 0 };
        byte[] png = AppleImageIOReference.EncodeGrayPng(pairs, 4, 1);
        _output.WriteLine(AppleImageIOReference.DescribeStoredImage(png));
        Assert.True(new OfficeImageIOCodec().TryDecode(png, null, out var image));
        byte[] actual = image!.GetPixels();
        for (int i = 0; i < 4; i++) {
            int expectedGray = pairs[i * 2 + 1] == 0 ? 0 : pairs[i * 2];
            for (int c = 0; c < 3; c++) Assert.InRange(Math.Abs(actual[i * 4 + c] - expectedGray), 0, 1);
            Assert.Equal(pairs[i * 2 + 1], actual[i * 4 + 3]);
        }
    }

    [SkippableFact]
    public void DimensionsAreRejectedBeforeRasterAllocation() {
        RequireMac();
        byte[] png = AppleImageIOReference.Encode(AppleImageIOReference.ColorfulRgba(), 32, 24, "public.png");
        Assert.False(new OfficeImageIOCodec(maximumDecodedPixels: 767).TryDecode(png, null, out var image));
        Assert.Null(image);
        // A valid PNG with a huge IHDR must be rejected by metadata policy, without decoding pixels.
        byte[] huge = (byte[])png.Clone();
        WriteBigEndian(huge, 16, 100_000);
        WriteBigEndian(huge, 20, 100_000);
        WriteBigEndian(huge, 29, Crc32(huge, 12, 17));
        Assert.False(new OfficeImageIOCodec().TryDecode(huge, null, out image));
        Assert.Null(image);
    }

    [SkippableTheory]
    [InlineData(false)]
    [InlineData(true)]
    public void OriginalHeicDecodesThroughTheSystemCodec(bool displayP3) {
        RequireMac();
        byte[] heic = AppleImageIOReference.CreateColorfulHeic(displayP3);
        _output.WriteLine(AppleImageIOReference.DescribeStoredImage(heic));
        Assert.True(new OfficeImageIOCodec().TryDecode(heic, "image/heic", out var image));
        Assert.Equal(32, image!.Width);
        Assert.Equal(24, image.Height);
        byte[] pixels = image.GetPixels();
        for (int i = 3; i < pixels.Length; i += 4) Assert.Equal(255, pixels[i]);
        byte[] reference = AppleImageIOReference.DecodeFullFile(heic).GetPixels();
        int maximumDifference = pixels.Zip(reference, (a, b) => Math.Abs(a - b)).Max();
        _output.WriteLine($"Raw vs native sRGB max channel difference: {maximumDifference}");
        if (displayP3) Assert.True(maximumDifference > 10);
        else Assert.InRange(maximumDifference, 0, 2);
    }

    [SkippableFact]
    public void FullFileReferenceAppliesOrientationWithoutUsingEmbeddedPreview() {
        RequireMac();
        byte[] expected = AppleImageIOReference.ColorfulRgba();
        byte[] png = AppleImageIOReference.Encode(expected, 32, 24, "public.png", orientation: 6);
        var image = AppleImageIOReference.DecodeFullFile(png);
        Assert.Equal(24, image.Width);
        Assert.Equal(32, image.Height);
        byte[] rotated = image.GetPixels();
        for (int y = 0; y < 24; y++) for (int x = 0; x < 32; x++) {
            int source = (y * 32 + x) * 4;
            int target = (x * 24 + (23 - y)) * 4;
            for (int c = 0; c < 4; c++) Assert.InRange(Math.Abs(expected[source + c] - rotated[target + c]), 0, 1);
        }
    }

    [SkippableFact]
    public void PublicTenBitP3FixtureHasAUsableNativeReference() {
        RequireMac();
        string path = Path.Combine(AppContext.BaseDirectory, "Images", "Heic", "display-p3-10bit.heic");
        Assert.True(File.Exists(path), $"Missing HEIC fixture: {path}");
        byte[] bytes = File.ReadAllBytes(path);
        _output.WriteLine(AppleImageIOReference.DescribeStoredImage(bytes));
        var reference = AppleImageIOReference.DecodeFullFile(bytes, maximumPixels: 16_000_000);
        Assert.True(reference.Width > 0 && reference.Height > 0);
        _output.WriteLine($"Native transformed sRGB reference: {reference.Width}x{reference.Height}");
    }

    [SkippableTheory]
    [InlineData("basic-8bit-srgb.heic")]
    [InlineData("display-p3-10bit.heic")]
    [InlineData("grid-12mp-512px-tiles.heic")]
    [InlineData("alpha-transparency.heic")]
    public void CoreSuppliedStrippedTileDecodesWithoutColorMatching(string fixture) {
        RequireMac();
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Images", "Heic", fixture));
        var probe = new TileProbe();
        // Capture the codec boundary; full-container comparison assertions are separate.
        bool composed = OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions { ImageCodec = probe }, out _, out var info);
        _output.WriteLine($"Core probe completed={composed}, diagnostic={info.Diagnostic}");
        Assert.NotNull(probe.Tile);
        Assert.True(probe.Decoded);
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(probe.Tile!, out var tileInfo));
        Assert.Single(tileInfo!.Items);
        Assert.Equal("hvc1", tileInfo.PrimaryItem!.ItemType);
        Assert.DoesNotContain(tileInfo.PrimaryItem.PropertyAssociations, p => p.PropertyType is "colr" or "irot" or "imir" or "clap");
        _output.WriteLine($"Original: {AppleImageIOReference.DescribeStoredImage(bytes)}");
        _output.WriteLine($"Stripped tile: {AppleImageIOReference.DescribeStoredImage(probe.Tile!)}");
    }

    private sealed class TileProbe : IOfficeRasterImageCodec {
        internal byte[]? Tile { get; private set; }
        internal bool Decoded { get; private set; }
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            bool success = new OfficeImageIOCodec().TryDecode(encodedBytes, contentType, out image);
            Tile ??= encodedBytes;
            Decoded = success;
            return success;
        }
    }

    [SkippableTheory]
    [InlineData("basic-8bit-srgb.heic")]
    [InlineData("generic-mif1-brand.heif")]
    [InlineData("orientation-2.heic")]
    [InlineData("orientation-6.heic")]
    [InlineData("orientation-8.heic")]
    [InlineData("display-p3-10bit.heic")]
    [InlineData("grid-12mp-512px-tiles.heic")]
    [InlineData("alpha-transparency.heic")]
    public void RealFixtureCompositionMatchesNativeFullFileReference(string fixture) {
        RequireMac();
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Images", "Heic", fixture));
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { ImageCodec = new OfficeImageIOCodec() }, out var actual, out var info), info.Diagnostic);
        var expected = AppleImageIOReference.DecodeFullFile(bytes, 16_000_000);
        AssertEquivalent(fixture, actual!, expected);
    }

    [SkippableFact]
    public void RealPhotoPickupMatchesNativeFullFileReferenceWhenProvided() {
        RequireMac();
        string? configured = Environment.GetEnvironmentVariable("OFFICEIMO_HEIC_REAL_PHOTOS");
        string directory = string.IsNullOrWhiteSpace(configured) ? FindRealPhotoDirectory() : configured;
        var files = Directory.Exists(directory)
            ? Directory.EnumerateFiles(directory, "*", new EnumerationOptions {
                RecurseSubdirectories = true, AttributesToSkip = FileAttributes.ReparsePoint
            }).Where(p => Path.GetExtension(p).ToLowerInvariant() is ".heic" or ".heif" or ".hif").OrderBy(p => p).ToArray()
            : Array.Empty<string>();
        Skip.If(files.Length == 0, "No real photos supplied. Set OFFICEIMO_HEIC_REAL_PHOTOS or use OfficeIMO.TestAssets/Images/Heic/RealPhotos.");
        var failures = new List<string>();
        foreach (string path in files) {
            string name = Path.GetFileName(path);
            try {
                if (new FileInfo(path).Length > 32 * 1024 * 1024)
                    throw new InvalidOperationException("Photo exceeds the native reference's 32 MiB input budget.");
                byte[] bytes = File.ReadAllBytes(path);
                _output.WriteLine($"{name}: {AppleImageIOReference.DescribeStoredImage(bytes)}");
                Assert.True(OfficeRasterImageDecoder.TryDecode(bytes,
                    new OfficeRasterDecodeOptions { ImageCodec = new OfficeImageIOCodec() }, out var actual, out var info), info.Diagnostic);
                var expected = AppleImageIOReference.DecodeFullFile(bytes, 16_000_000);
                AssertEquivalent(name, actual!, expected);
            } catch (Exception error) {
                failures.Add($"{name}: {error.Message}");
                _output.WriteLine($"Mismatch: {failures[failures.Count - 1]}");
            }
        }
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
    }

    private static string FindRealPhotoDirectory() {
        for (DirectoryInfo? directory = new DirectoryInfo(AppContext.BaseDirectory); directory != null; directory = directory.Parent) {
            string assets = Path.Combine(directory.FullName, "OfficeIMO.TestAssets", "Images", "Heic");
            if (Directory.Exists(assets)) return Path.Combine(assets, "RealPhotos");
        }
        return Path.Combine(AppContext.BaseDirectory, "Images", "Heic", "RealPhotos");
    }

    private void AssertEquivalent(string fixture, OfficeRasterImage actual, OfficeRasterImage expected) {
        _output.WriteLine($"{fixture}: Core={actual.Width}x{actual.Height}; native={expected.Width}x{expected.Height}");
        Assert.Equal(expected.Width, actual.Width);
        Assert.Equal(expected.Height, actual.Height);
        byte[] a = actual.GetPixels(), b = expected.GetPixels();
        long colorError = 0, alphaError = 0, premultipliedError = 0, colors = 0;
        int maxColor = 0, maxAlpha = 0, outlierPixels = 0;
        var examples = new List<string>();
        for (int i = 0; i < a.Length; i += 4) {
            int alpha = Math.Abs(a[i + 3] - b[i + 3]);
            alphaError += alpha;
            maxAlpha = Math.Max(maxAlpha, alpha);
            for (int c = 0; c < 3; c++)
                premultipliedError += Math.Abs((a[i + c] * a[i + 3] + 127) / 255 - (b[i + c] * b[i + 3] + 127) / 255);
            // RGB under near-transparent pixels is ill-conditioned after 8-bit unpremultiplication.
            if (a[i + 3] < 250 || b[i + 3] < 250) continue;
            int pixelMaximum = 0;
            for (int c = 0; c < 3; c++) {
                int difference = Math.Abs(a[i + c] - b[i + c]);
                colorError += difference; colors++; maxColor = Math.Max(maxColor, difference);
                pixelMaximum = Math.Max(pixelMaximum, difference);
            }
            if (pixelMaximum > 16) {
                outlierPixels++;
                if (examples.Count < 8) examples.Add($"({i / 4 % actual.Width},{i / 4 / actual.Width}):{pixelMaximum}");
            }
        }
        double meanColor = colorError / (double)Math.Max(1, colors), meanAlpha = alphaError / (double)(a.Length / 4);
        double meanPremultiplied = premultipliedError / (double)(a.Length / 4 * 3);
        _output.WriteLine($"Opaque RGB mean={meanColor:F4}, max={maxColor}; alpha mean={meanAlpha:F4}, max={maxAlpha}");
        _output.WriteLine($"Premultiplied RGB mean across all pixels: {meanPremultiplied:F4}");
        _output.WriteLine($"RGB pixels with channel error >16: {outlierPixels}; first coordinates: {string.Join(", ", examples)}");
        Assert.True(meanColor <= 3, $"Mean opaque RGB difference {meanColor:F4} exceeds 3.");
        Assert.True(meanPremultiplied <= 3, $"Mean premultiplied RGB difference {meanPremultiplied:F4} exceeds 3.");
        Assert.True(meanAlpha <= 1, $"Mean alpha difference {meanAlpha:F4} exceeds 1.");
    }

    private static void RequireMac() => Skip.IfNot(OperatingSystem.IsMacOSVersionAtLeast(10, 13), "Requires macOS ImageIO.");
    private static void WriteBigEndian(byte[] bytes, int offset, uint value) {
        for (int i = 3; i >= 0; i--) { bytes[offset + i] = (byte)value; value >>= 8; }
    }
    private static uint Crc32(byte[] bytes, int offset, int count) {
        uint crc = uint.MaxValue;
        for (int i = offset; i < offset + count; i++) {
            crc ^= bytes[i];
            for (int j = 0; j < 8; j++) crc = (crc >> 1) ^ ((crc & 1) == 0 ? 0 : 0xEDB88320);
        }
        return ~crc;
    }
}
