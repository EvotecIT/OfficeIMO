using System;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class DrawingRasterEncodingTests {
    [Theory]
    [InlineData(5)]
    [InlineData(9)]
    public void OptimalPngObservesCancellationInLaterCompressionPasses(int cancelAtRow) {
        var image = new OfficeRasterImage(32, 4, OfficeColor.CornflowerBlue);
        using var cancellation = new CancellationTokenSource();
        using var destination = new MemoryStream();
        int rows = 0;
        Assert.Throws<OperationCanceledException>(() =>
            OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Png, destination,
                new OfficeRasterEncodingOptions(), maximumEncodedBytes: long.MaxValue,
                cancellationToken: cancellation.Token, checkpointObserver: checkpoint => {
                    if (checkpoint == OfficeRasterEncodingCheckpoint.PngCompressionRow && ++rows == cancelAtRow)
                        cancellation.Cancel();
                }));
        Assert.Equal(cancelAtRow, rows);
        destination.WriteByte(0x7A);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OptimalPngFitsTheUnfilteredLosslessRepresentation(bool adaptivePreferred) {
        const int width = 160, height = 80;
        var image = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                int value = adaptivePreferred
                    ? (((x / 16 + y / 16) & 1) != 0 ? x * 19 % 256 : 255)
                    : (x % 19 < 10 && y % 13 < 7 ? 17 * ((x + y) % 16) : 255);
                image.SetPixel(x, y, OfficeColor.FromRgb((byte)value, (byte)value, (byte)value));
            }
        }

        // PNG filter zero carries the unchanged RGBA row. This independent
        // representation catches a filter heuristic that expands repeated edges.
        byte[] pixels = image.GetPixels();
        int stride = width * 4;
        var rows = new byte[(stride + 1) * height];
        for (int y = 0; y < height; y++) Buffer.BlockCopy(pixels, y * stride, rows, y * (stride + 1) + 1, stride);
        byte[] unfiltered = OfficePngWriter.EncodeScanlines(width, height, 8, 6, rows);
        byte[] optimal = OfficePngWriter.Encode(image);
        using var stream = new MemoryStream();
        OfficePngWriter.EncodeTo(image, stream);

        Assert.True(optimal.Length <= unfiltered.Length);
        Assert.True(stream.Length <= unfiltered.Length);
        if (adaptivePreferred) Assert.True(optimal.Length < unfiltered.Length);
        Assert.True(OfficePngReader.TryDecode(optimal, out OfficeRasterImage? decoded));
        Assert.Equal(pixels, decoded!.GetPixels());
        Assert.True(OfficePngReader.TryDecode(stream.ToArray(), out decoded));
        Assert.Equal(pixels, decoded!.GetPixels());
    }
}
