using System;
using System.IO;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Drawing.Wic.Tests;

public sealed class HeicCompositionIntegrationTests {
    private readonly ITestOutputHelper _output;

    public HeicCompositionIntegrationTests(ITestOutputHelper output) => _output = output;

    [WindowsHeicFact]
    public void CoreComposition_DecodesCc0CorpusThroughSanitizedWicTiles() {
        if (!OperatingSystem.IsWindows()) return;
        // These are independent encoded inputs, not publisher-generated reference rasters.
        // Core owns grids, transforms and profiles; WIC receives only its sanitized hvc1 tiles.
        (string File, int Width, int Height)[] fixtures = {
            ("basic-8bit-srgb.heic", 1600, 1200),
            ("generic-mif1-brand.heif", 1600, 1200),
            ("grid-12mp-512px-tiles.heic", 4032, 3024),
            ("display-p3-10bit.heic", 1600, 1200),
            ("alpha-transparency.heic", 1200, 1200),
            ("orientation-2.heic", 1200, 800),
            ("orientation-6.heic", 1200, 800),
            ("orientation-8.heic", 1200, 800)
        };
        foreach (var fixture in fixtures) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Heic", fixture.File));
            Assert.True(OfficeRasterImageDecoder.TryDecode(bytes,
                new OfficeRasterDecodeOptions { ImageCodec = new OfficeWicCodec() },
                out OfficeRasterImage? image, out OfficeRasterDecodeInfo info),
                $"{fixture.File}: {info.Failure} ({info.CodecUnavailableReason})");
            Assert.NotNull(image);
            Assert.Equal(fixture.Width, image.Width);
            Assert.Equal(fixture.Height, image.Height);
            Assert.Equal(OfficeImageFormat.Heic, info.Format);
            Assert.True(WindowsWicReference.TryDecode(bytes, out WindowsWicReference? reference, out string detail),
                $"{fixture.File}: {detail}");
            Assert.NotNull(reference);
            _output.WriteLine($"{fixture.File}: {detail}");
            WicImageComparison.AssertEquivalent(fixture.File, image, reference, _output);
        }
    }
}
