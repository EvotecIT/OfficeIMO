using System;
using System.IO;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ImageEncodingEffortTests {
    // Each effort selects a different encoder path. Protect exact hidden RGB,
    // deterministic byte/stream output and option cloning through conversion APIs.
    [Theory]
    [InlineData(OfficeImageExportFormat.Png, OfficeImageEncodingEffort.Fast)]
    [InlineData(OfficeImageExportFormat.Png, OfficeImageEncodingEffort.Balanced)]
    [InlineData(OfficeImageExportFormat.Png, OfficeImageEncodingEffort.Smallest)]
    [InlineData(OfficeImageExportFormat.Webp, OfficeImageEncodingEffort.Fast)]
    [InlineData(OfficeImageExportFormat.Webp, OfficeImageEncodingEffort.Balanced)]
    [InlineData(OfficeImageExportFormat.Webp, OfficeImageEncodingEffort.Smallest)]
    public void EffortSurvivesConversionOptionsAndPreservesDeterministicSamples(
        OfficeImageExportFormat format, OfficeImageEncodingEffort effort) {
        var options = new OfficeImageExportOptions();
        options.RasterEncoding.Png.Effort = effort;
        options.RasterEncoding.Webp.Effort = effort;
        var effective = options.CreateEffectiveImageExportOptions<OfficeImageExportOptions>();
        Assert.NotSame(options.RasterEncoding.Png, effective.RasterEncoding.Png);
        Assert.NotSame(options.RasterEncoding.Webp, effective.RasterEncoding.Webp);
        Assert.Equal(effort, effective.RasterEncoding.Png.Effort);
        Assert.Equal(effort, effective.RasterEncoding.Webp.Effort);
        foreach (bool palette in new[] { false, true }) {
            var pixels = new byte[193 * 65 * 4];
            var random = new Random(3901);
            for (int i = 0; i < pixels.Length; i += 4) {
                int color = palette ? random.Next(16) : random.Next(256);
                pixels[i] = (byte)color;
                pixels[i + 1] = (byte)(color * 71);
                pixels[i + 2] = (byte)(color * 193);
                pixels[i + 3] = palette ? (byte)(color % 3 == 0 ? 0 : 255) : (byte)random.Next(256);
            }
            var image = OfficeRasterImage.FromRgba32(193, 65, pixels);
            byte[] encoded = OfficeRasterImageEncoder.Encode(image, format, effective.RasterEncoding);
            Assert.Equal(encoded, OfficeRasterImageEncoder.Encode(image, format, effective.RasterEncoding));
            using var stream = new MemoryStream();
            OfficeRasterImageEncoder.EncodeTo(image, format, stream, effective.RasterEncoding);
            Assert.Equal(encoded, stream.ToArray());
            Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out var decoded));
            Assert.Equal(pixels, decoded!.GetPixels());
            Assert.Equal(pixels, image.GetPixels());
        }
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void DefaultIsBalancedAndInvalidEffortIsRejectedBeforeWriting(OfficeImageExportFormat format) {
        var image = new OfficeRasterImage(33, 17);
        var explicitOptions = new OfficeRasterEncodingOptions();
        Assert.Equal(OfficeImageEncodingEffort.Balanced, explicitOptions.Png.Effort);
        Assert.Equal(OfficeImageEncodingEffort.Balanced, explicitOptions.Webp.Effort);
        Assert.Equal(OfficeRasterImageEncoder.Encode(image, format), OfficeRasterImageEncoder.Encode(image, format, explicitOptions));
        explicitOptions.Png.Effort = explicitOptions.Webp.Effort = (OfficeImageEncodingEffort)99;
        using var output = new MemoryStream();
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterImageEncoder.EncodeTo(image, format, output, explicitOptions));
        Assert.Equal(0, output.Length);
    }
}
