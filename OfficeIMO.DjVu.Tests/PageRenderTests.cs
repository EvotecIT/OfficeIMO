using OfficeIMO.Drawing;
using System.Text.RegularExpressions;

namespace OfficeIMO.DjVu.Tests;

public sealed class PageRenderTests {
    [Theory]
    [InlineData("gradient")]
    [InlineData("palette")]
    [InlineData("standalone-color")]
    [InlineData("noise-full")]
    [InlineData("noise-normal")]
    [InlineData("gray")]
    [InlineData("sampling")]
    public void CompletePageMatchesIndependentNativeColorRaster(string name) {
        var page = Assert.Single(DjVuDocument.Load(Fixture(name + ".djvu")).Pages);
        var result = page.Render();
        var expected = ReadPpm(Fixture(name + "-reference.ppm"));
        Assert.Equal((expected.Width, expected.Height), (result.Image.Width, result.Image.Height));
        Assert.Equal(expected.GetPixels(), result.Image.GetPixels());
        Assert.Equal(page.Dpi, result.Dpi);
    }

    [Fact]
    public void NativeRegionUsesBottomLeftCoordinatesAndReturnedPixelsAreOwned() {
        var page = Assert.Single(DjVuDocument.Load(Fixture("palette.djvu")).Pages);
        var expected = ReadPpm(Fixture("palette-reference.ppm"));
        var settings = new DjVuRenderOptions { Region = new DjVuRectangle(3, 4, 35, 27) };
        var result = page.Render(settings);
        Assert.Equal((35, 27), (result.Image.Width, result.Image.Height));
        for (int y = 0; y < 27; y++) for (int x = 0; x < 35; x++) Assert.Equal(expected.GetPixel(x + 3, expected.Height - 4 - 27 + y), result.Image.GetPixel(x, y));
        result.Image.SetPixel(0, 0, OfficeColor.Blue);
        Assert.Equal(expected.GetPixels(), page.Render().Image.GetPixels());
        var scaled = page.Render(new DjVuRenderOptions { Dpi = 150, Region = settings.Region });
        Assert.Equal((18, 14), (scaled.Image.Width, scaled.Image.Height));
        Assert.Equal(150, scaled.Dpi);
    }

    [Fact]
    public void RotationMatchesIndependentDisplayRasterAndCanBeDisabled() {
        var page = Assert.Single(DjVuDocument.Load(Fixture("palette-rotated.djvu")).Pages);
        Assert.Equal(90, page.Rotation);
        var expected = ReadPpm(Fixture("palette-rotated-reference.ppm"));
        var rotated = page.Render();
        Assert.Equal((96, 128), (rotated.Image.Width, rotated.Image.Height));
        Assert.Equal(expected.GetPixels(), rotated.Image.GetPixels());
        Assert.Equal(ReadPpm(Fixture("palette-reference.ppm")).GetPixels(), page.Render(new DjVuRenderOptions { ApplyRotation = false }).Image.GetPixels());
    }

    [Theory]
    [InlineData("fax")]
    [InlineData("fax-striped")]
    [InlineData("fax-inverted")]
    public void FaxMasksReuseOwnedCoreCodecAndMatchIndependentRaster(string name) {
        var page = Assert.Single(DjVuDocument.Load(Fixture(name + ".djvu")).Pages);
        Assert.Equal(ReadPpm(Fixture(name + "-reference.ppm")).GetPixels(), page.Render().Image.GetPixels());
    }

    [Fact]
    public void RenderingRejectsExcessiveOutputsRegionsAndCancellation() {
        var page = Assert.Single(DjVuDocument.Load(Fixture("gradient.djvu")).Pages);
        Assert.Equal(nameof(DjVuRenderOptions.MaxPixels), Assert.Throws<DjVuResourceLimitException>(() => page.Render(new DjVuRenderOptions { MaxPixels = 10 })).LimitName);
        Assert.Throws<DjVuResourceLimitException>(() => page.Render(new DjVuRenderOptions { MaxBytes = 1024 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => page.Render(new DjVuRenderOptions { Region = new DjVuRectangle(64, 0, 2, 1) }));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => page.Render(cancellationToken: cancellation.Token));
    }

    [Fact]
    public void JpegLayerUsesManagedCoreAndIndependentReference() {
        var page = Assert.Single(DjVuDocument.Load(Fixture("jpeg-background.djvu")).Pages);
        var expected = ReadPpm(Fixture("jpeg-background-reference.ppm"));
        var actual = page.Render().Image;
        Assert.Equal((expected.Width, expected.Height), (actual.Width, actual.Height));
        byte[] native = expected.GetPixels(), owned = actual.GetPixels();
        Assert.True(native.Zip(owned, (a, b) => Math.Abs(a - b)).Max() <= 2);
        Assert.Throws<DjVuResourceLimitException>(() => page.Render(new DjVuRenderOptions { MaxBytes = 70_000 }));
    }

    [Fact]
    public void ShortIw44EdgesRetainMeasuredQualificationAndLosslessGate() {
        var page = Assert.Single(DjVuDocument.Load(Fixture("noise-small.djvu")).Pages);
        var rendered = page.Render();
        var expected = ReadPpm(Fixture("noise-small-reference.ppm"));
        byte[] native = expected.GetPixels(), actual = rendered.Image.GetPixels();
        int[] errors = native.Zip(actual, (a, b) => Math.Abs(a - b)).ToArray();
        Assert.True(errors.Max() <= 5);
        Assert.True(errors.Average() <= 0.25); // Includes the exact alpha channel.
        Assert.Contains(rendered.FidelityDiagnostics, d => d.Code == "djvu.render.short-iw44-edge" && d.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<OfficeConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void UnsupportedLayersAndExcessiveJpegDimensionsFailBeforePainting() {
        byte[] gradient = File.ReadAllBytes(Fixture("gradient.djvu"));
        var unsupported = DjVuDocument.Load(SharedComponentTests.AppendChunk(gradient, "BG2k", Array.Empty<byte>()));
        Assert.Throws<NotSupportedException>(() => unsupported.Pages[0].Render());
        byte[] jpeg = File.ReadAllBytes(Fixture("jpeg-background.djvu"));
        var original = DjVuDocument.Load(jpeg);
        var layer = original.PageChunks(original.Pages[0].Component, default).Single(c => c.Id == "BGjp");
        int sof = -1;
        for (int i = layer.Offset; i < layer.Offset + layer.Length - 8; i++) if (jpeg[i] == 255 && jpeg[i + 1] == 192) { sof = i; break; }
        Assert.True(sof >= 0);
        jpeg[sof + 5] = jpeg[sof + 6] = jpeg[sof + 7] = jpeg[sof + 8] = 255;
        Assert.Equal(nameof(DjVuReadOptions.MaxPagePixels), Assert.Throws<DjVuResourceLimitException>(() => DjVuDocument.Load(jpeg).Pages[0].Render()).LimitName);
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Authored", name);
    internal static OfficeRasterImage ReadPpm(string path) {
        byte[] bytes = File.ReadAllBytes(path);
        var header = Regex.Match(Encoding.ASCII.GetString(bytes, 0, Math.Min(64, bytes.Length)), @"^P6\s+(\d+)\s+(\d+)\s+255\s");
        Assert.True(header.Success);
        int width = int.Parse(header.Groups[1].Value), height = int.Parse(header.Groups[2].Value);
        var rgba = new byte[width * height * 4];
        Assert.Equal(width * height * 3, bytes.Length - header.Length);
        for (int i = 0, j = header.Length; i < rgba.Length; i += 4, j += 3) { rgba[i] = bytes[j]; rgba[i + 1] = bytes[j + 1]; rgba[i + 2] = bytes[j + 2]; rgba[i + 3] = 255; }
        return OfficeRasterImage.FromRgba32(width, height, rgba);
    }
}
