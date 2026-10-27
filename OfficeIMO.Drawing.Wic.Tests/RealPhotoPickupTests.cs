using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Drawing.Wic.Tests;

public sealed class RealPhotoPickupTests {
    private readonly ITestOutputHelper _output;

    public RealPhotoPickupTests(ITestOutputHelper output) => _output = output;

    [SkippableFact]
    public void RealPhotoPickupMatchesNativeFullFileReferenceWhenProvided() {
        Skip.IfNot(OperatingSystem.IsWindows(), "WIC camera-photo qualification requires Windows.");
        if (!OperatingSystem.IsWindows()) return;
        OfficeRasterCodecAvailability availability = new OfficeWicCodec().GetAvailability();
        Skip.IfNot(availability.IsAvailable, availability.Message);
        string? configured = Environment.GetEnvironmentVariable("OFFICEIMO_HEIC_REAL_PHOTOS");
        string directory = string.IsNullOrWhiteSpace(configured) ? FindRealPhotoDirectory() : configured;
        string[] files = Directory.Exists(directory)
            ? Directory.EnumerateFiles(directory, "*", new EnumerationOptions {
                RecurseSubdirectories = true, AttributesToSkip = FileAttributes.ReparsePoint
            }).Where(p => Path.GetExtension(p).ToLowerInvariant() is ".heic" or ".heif" or ".hif")
                .OrderBy(p => p, StringComparer.Ordinal).ToArray()
            : Array.Empty<string>();
        Skip.If(files.Length == 0,
            "No real photos supplied. Set OFFICEIMO_HEIC_REAL_PHOTOS or use OfficeIMO.TestAssets/Images/Heic/RealPhotos.");

        var failures = new List<string>();
        int compared = 0;
        foreach (string path in files) {
            string name = Path.GetFileName(path);
            try {
                if (new FileInfo(path).Length > 32 * 1024 * 1024)
                    throw new InvalidOperationException("Photo exceeds the full-file reference's 32 MiB input budget.");
                byte[] bytes = File.ReadAllBytes(path);
                if (!WindowsWicReference.TryDecode(bytes, out WindowsWicReference? reference, out string detail)) {
                    _output.WriteLine($"{name}: comparison unavailable: {detail}");
                    continue;
                }
                _output.WriteLine($"{name}: {detail}");
                Assert.True(OfficeRasterImageDecoder.TryDecode(bytes,
                    new OfficeRasterDecodeOptions { ImageCodec = new OfficeWicCodec() },
                    out OfficeRasterImage? actual, out OfficeRasterDecodeInfo info),
                    $"{name}: Core decode failed: {info.Failure} ({info.CodecUnavailableReason}).");
                Assert.NotNull(actual);
                Assert.NotNull(reference);
                WicImageComparison.AssertEquivalent(name, actual, reference, _output);
                compared++;
            } catch (Exception ex) {
                failures.Add($"{name}: {ex.Message}");
                _output.WriteLine(failures[^1]);
            }
        }
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
        Skip.If(compared == 0, "No supplied photo had an available bounded native full-file WIC reference; see test output.");
        _output.WriteLine($"Compared {compared}/{files.Length} real photos; {files.Length - compared} native references unavailable.");
    }

    private static string FindRealPhotoDirectory([CallerFilePath] string sourceFile = "") {
        // Test outputs may be outside the checkout via --artifacts-path.
        string sourceAssets = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(sourceFile)!,
            "..", "OfficeIMO.TestAssets", "Images", "Heic", "RealPhotos"));
        if (Directory.Exists(sourceAssets)) return sourceAssets;
        for (DirectoryInfo? directory = new(AppContext.BaseDirectory); directory != null; directory = directory.Parent) {
            string assets = Path.Combine(directory.FullName, "OfficeIMO.TestAssets", "Images", "Heic", "RealPhotos");
            if (Directory.Exists(assets)) return assets;
        }
        return Path.Combine(AppContext.BaseDirectory, "Images", "Heic", "RealPhotos");
    }

}

[System.Runtime.Versioning.SupportedOSPlatform("windows")]
internal static class WicImageComparison {
    internal static void AssertEquivalent(string name, OfficeRasterImage actual, WindowsWicReference expected, ITestOutputHelper output) {
        output.WriteLine($"{name}: Core={actual.Width}x{actual.Height}; native={expected.Width}x{expected.Height}");
        Assert.Equal(expected.Width, actual.Width);
        Assert.Equal(expected.Height, actual.Height);
        long opaqueError = 0, opaqueSamples = 0, premultipliedError = 0, alphaError = 0;
        int maxColor = 0, maxAlpha = 0, outliers = 0;
        var examples = new List<string>();
        for (int y = 0; y < actual.Height; y++) {
            for (int x = 0; x < actual.Width; x++) {
                OfficeColor a = actual.GetPixel(x, y), b = expected.GetPixel(x, y);
                int alpha = Math.Abs(a.A - b.A);
                alphaError += alpha;
                maxAlpha = Math.Max(maxAlpha, alpha);
                premultipliedError += PremultipliedError(a.R, a.A, b.R, b.A) +
                    PremultipliedError(a.G, a.A, b.G, b.A) + PremultipliedError(a.B, a.A, b.B, b.A);
                if (a.A < 250 || b.A < 250) continue;
                int r = Math.Abs(a.R - b.R), g = Math.Abs(a.G - b.G), blue = Math.Abs(a.B - b.B);
                int maximum = Math.Max(r, Math.Max(g, blue));
                opaqueError += r + g + blue;
                opaqueSamples += 3;
                maxColor = Math.Max(maxColor, maximum);
                if (maximum > 16) {
                    outliers++;
                    if (examples.Count < 8) examples.Add($"({x},{y}):{maximum}");
                }
            }
        }
        long pixels = (long)actual.Width * actual.Height;
        double meanOpaque = opaqueError / (double)Math.Max(1, opaqueSamples);
        double meanPremultiplied = premultipliedError / (double)(pixels * 3);
        double meanAlpha = alphaError / (double)pixels;
        output.WriteLine($"Opaque RGB mean={meanOpaque:F4}, max={maxColor}; alpha mean={meanAlpha:F4}, max={maxAlpha}");
        output.WriteLine($"Premultiplied RGB mean={meanPremultiplied:F4}; RGB pixels with channel error >16: {outliers}; first: {string.Join(", ", examples)}");
        Assert.True(meanOpaque <= 3, $"Mean opaque RGB difference {meanOpaque:F4} exceeds 3.");
        Assert.True(meanPremultiplied <= 3, $"Mean premultiplied RGB difference {meanPremultiplied:F4} exceeds 3.");
        Assert.True(meanAlpha <= 1, $"Mean alpha difference {meanAlpha:F4} exceeds 1.");
    }

    private static int PremultipliedError(byte a, byte alphaA, byte b, byte alphaB) =>
        Math.Abs((a * alphaA + 127) / 255 - (b * alphaB + 127) / 255);
}
