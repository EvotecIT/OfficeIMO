using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLosslessJpegEmbeddingTests {
    [Theory]
    [InlineData("Lossless8.jpg")]
    [InlineData("Lossless16.jpg")]
    public void LosslessJpegIsNormalizedBeforePdfEmbedding(string file) {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", file));
        Assert.True(PdfDocument.TryPrepareImageBytes(jpeg, out byte[] normalized, out var info,
            out bool transcoded, out string? reason), reason);
        Assert.True(transcoded);
        Assert.Equal(OfficeImageFormat.Png, info!.Format);
        byte[] pdf = PdfDocument.Create().Image(jpeg, 24, 24).ToBytes();
        var embedded = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        Assert.Equal("image/png", embedded.MimeType);
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var expected));
        Assert.True(OfficePngReader.TryDecode(embedded.Bytes, out var actual));
        Assert.Equal((expected!.Width, expected.Height), (actual!.Width, actual.Height));
        for (int y = 0; y < expected.Height; y++) for (int x = 0; x < expected.Width; x++)
            Assert.Equal(expected.GetPixel(x, y), actual.GetPixel(x, y));
    }
}
