using System;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterEncodedInputTests {
    [Fact]
    public void EncodedInputOwnsItsBytesAndPreservesTheCallerStreamPosition() {
        byte[] source = { 1, 2, 3, 4 };
        using var stream = new MemoryStream(source, 0, source.Length, writable: true, publiclyVisible: true);
        byte[] bytes = OfficeRasterImageDecoder.ReadEncodedBytes(stream);
        Assert.Equal(source, bytes);
        Assert.Equal(0L, stream.Position);
        source[0] = 9;
        Assert.Equal(1, bytes[0]);
        stream.Position = 2;
        Assert.Equal(new byte[] { 3, 4 }, OfficeRasterImageDecoder.ReadEncodedBytes(stream));
        Assert.Equal(2L, stream.Position);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void EncodedInputRejectsOverBudgetAndEmptyStreamsAndPreservesPosition() {
        using var stream = new MemoryStream(new byte[] { 1, 2, 3 });
        var options = new OfficeRasterDecodeOptions { MaximumEncodedBytes = 2 };
        Assert.Throws<InvalidDataException>(() => OfficeRasterImageDecoder.ReadEncodedBytes(stream, options));
        Assert.Equal(0L, stream.Position);
        stream.Position = stream.Length;
        Assert.Throws<InvalidDataException>(() => OfficeRasterImageDecoder.ReadEncodedBytes(stream));
        Assert.Equal(stream.Length, stream.Position);
    }

    [Fact]
    public void EncodedInputObservesCancellationWithoutConsumingASeekableStream() {
        using var stream = new MemoryStream(new byte[] { 1, 2, 3 });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.ReadEncodedBytes(stream,
            new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token }));
        Assert.Equal(0L, stream.Position);
    }
}
