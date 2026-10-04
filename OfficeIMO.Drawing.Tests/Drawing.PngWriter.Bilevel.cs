using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingPngBilevelWriterTests {
    [Theory]
    [InlineData(1)]
    [InlineData(7)]
    [InlineData(8)]
    [InlineData(9)]
    [InlineData(1025)]
    [InlineData(16385)]
    public void OptimalBilevelPngPreservesBitAndRowTailsAcrossEncodingSurfaces(int width) {
        var image = CreateScan(width, 4);
        var options = new OfficePngEncodeOptions { DpiX = 300, DpiY = 144 };
        byte[] png = OfficePngWriter.Encode(image, options);
        using var stream = new MemoryStream();
        OfficePngWriter.EncodeTo(image, stream, options);
        Assert.Equal(png, stream.ToArray());
        Assert.Equal(png, OfficePngWriter.EncodeRgba(width, image.Height, image.GetPixels(), options));
#if NET8_0_OR_GREATER
        var writer = new ArrayBufferWriter<byte>();
        OfficePngWriter.EncodeTo(image, writer, options);
        Assert.Equal(png, writer.WrittenSpan.ToArray());
#endif
        Assert.Equal(1, png[24]);
        Assert.Equal(0, png[25]);
        Assert.True(OfficePngReader.TryDecode(png, out var decoded));
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
        var info = OfficeImageReader.Identify(png);
        Assert.InRange(info.DpiX, 299.98, 300.02);
        Assert.InRange(info.DpiY, 143.98, 144.02);

        byte[] noMetadata = OfficePngWriter.Encode(image, CancellationToken.None);
        Assert.True(OfficePngReader.TryDecode(noMetadata, out decoded));
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
        byte[] stored = OfficePngWriter.Encode(image, OfficePngCompression.Stored);
        Assert.Equal(8, stored[24]);
        Assert.Equal(6, stored[25]);
        Assert.True(OfficePngReader.TryDecode(stored, out decoded));
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
    }

    [Theory]
    [InlineData(128, 128, 128, 255)]
    [InlineData(0, 1, 0, 255)]
    [InlineData(255, 255, 255, 254)]
    [InlineData(17, 31, 255, 0)]
    public void AFinalNonBilevelPixelPreservesEveryRgbaChannel(byte r, byte g, byte b, byte alpha) {
        var image = CreateScan(1025, 4);
        image.SetPixel(image.Width - 1, image.Height - 1, OfficeColor.FromRgba(r, g, b, alpha));
        byte[] png = OfficePngWriter.Encode(image);
        Assert.Equal(8, png[24]);
        Assert.Equal(6, png[25]);
        Assert.True(OfficePngReader.TryDecode(png, out var decoded));
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
    }

    [Fact]
    public void BilevelSelectionObservesCancellationBeforeWritingTheDestination() {
        var image = CreateScan(8193, 2);
        using var destination = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        int selections = 0;
        var exception = Assert.Throws<OperationCanceledException>(() =>
            OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Png, destination,
                new OfficeRasterEncodingOptions(), long.MaxValue, cancellation.Token, checkpoint => {
                    if (checkpoint == OfficeRasterEncodingCheckpoint.PngColorSelectionBlock && ++selections == 2)
                        cancellation.Cancel();
                }));
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(0, destination.Length);
        destination.WriteByte(0x7A);
    }

    [Theory]
    [InlineData(1, 2)]
    [InlineData(2, 1)]
    [InlineData(5, 1)]
    [InlineData(9, 1)]
    public void BilevelPackingAndAllCompressionPassesRemainCancellable(int cancelAtRow, int cancelAtBlock) {
        var image = CreateScan(8193, 4);
        using var destination = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        int rows = 0, blocks = 0;
        var exception = Assert.Throws<OperationCanceledException>(() =>
            OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Png, destination,
                new OfficeRasterEncodingOptions(), long.MaxValue, cancellation.Token, checkpoint => {
                    if (checkpoint == OfficeRasterEncodingCheckpoint.PngCompressionRow) { rows++; blocks = 0; }
                    if (rows == cancelAtRow && checkpoint == OfficeRasterEncodingCheckpoint.PngFilteringBlock
                        && ++blocks == cancelAtBlock) cancellation.Cancel();
                }));
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(cancelAtRow, rows);
        Assert.Equal(cancelAtBlock, blocks);
        destination.WriteByte(0x7A);
    }

    [Fact]
    public void BilevelOutputHonorsTheExactEncodedByteBudget() {
        var image = CreateScan(137, 20);
        var options = new OfficeRasterEncodingOptions { WriteResolutionMetadata = false };
        byte[] expected = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png, options);
        Assert.Equal(expected, OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png, options, expected.Length));
        Assert.Throws<OfficeImageExportBatchLimitException>(() =>
            OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png, options, expected.Length - 1));
        using var destination = new MemoryStream();
        Assert.Throws<OfficeImageExportBatchLimitException>(() => OfficeRasterImageEncoder.EncodeTo(
            image, OfficeImageExportFormat.Png, destination, options, expected.Length - 1));
        Assert.True(destination.Length <= expected.Length - 1);
        destination.WriteByte(0x7A);
    }

    private static OfficeRasterImage CreateScan(int width, int height) {
        var image = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++) for (int x = 0; x < width; x++)
            image.SetPixel(x, y, ((x * 13 + y * 7) & 3) == 0 ? OfficeColor.White : OfficeColor.Black);
        return image;
    }
}
