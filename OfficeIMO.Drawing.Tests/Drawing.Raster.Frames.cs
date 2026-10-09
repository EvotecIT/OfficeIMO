using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingRasterTests {
    [Fact]
    public void AllFrameDecodeReturnsIndependentRenderedBuffersAndRetainsEveryFrame() {
        Assert.True(OfficeRasterImageDecoder.TryDecodeFrames(CreateTwoFrameGif(), null, out OfficeRasterFrames? frames));
        Assert.Equal(2, frames!.Count);
        Assert.Equal(OfficeColor.Red, frames[0].Image.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Lime, frames[1].Image.GetPixel(0, 0));
        frames[0].Image.SetPixel(0, 0, OfficeColor.Blue);
        Assert.Equal(OfficeColor.Lime, frames[1].Image.GetPixel(0, 0));
        Assert.Equal(1, frames.PlayCount);
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(1, 2)]
    [InlineData(3, 4)]
    public void AllFrameDecodeNormalizesGifRepeatsAndRetainsTiming(int repeats, int totalPlays) {
        byte[] gif = CreateSinglePixelGif();
        int descriptor = Array.IndexOf(gif, (byte)0x2C);
        byte[] extension = {
            0x21, 0xFF, 0x0B,
            (byte)'N', (byte)'E', (byte)'T', (byte)'S', (byte)'C', (byte)'A',
            (byte)'P', (byte)'E', (byte)'2', (byte)'.', (byte)'0',
            0x03, 0x01, (byte)repeats, 0x00, 0x00,
            0x21, 0xF9, 0x04, 0x00, 0x0A, 0x00, 0x00, 0x00
        };
        byte[] animated = gif.Take(descriptor).Concat(extension).Concat(gif.Skip(descriptor)).ToArray();

        Assert.True(OfficeRasterImageDecoder.TryDecodeFrames(animated, null, out OfficeRasterFrames? frames));
        Assert.Single(frames!);
        Assert.Equal(totalPlays, frames!.PlayCount);
        Assert.Equal(TimeSpan.FromMilliseconds(100), frames[0].Duration);
        Assert.Equal(OfficeColor.White, frames[0].Image.GetPixel(0, 0));
    }

    [Fact]
    public void AllFrameDecodeRejectsIncompleteOrOverBudgetSequencesWithoutPartialResults() {
        byte[] gif = CreateTwoFrameGif();
        Assert.False(OfficeRasterImageDecoder.TryDecodeFrames(gif, null, out var frames, maximumFrames: 1));
        Assert.Null(frames);
        Assert.False(OfficeRasterImageDecoder.TryDecodeFrames(gif,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 1 }, out frames));
        Assert.Null(frames);
        Assert.False(OfficeRasterImageDecoder.TryDecodeFrames(gif.Take(gif.Length - 1).ToArray(), null, out frames));
        Assert.Null(frames);
    }

    [Fact]
    public void AllFrameDecodeObservesCancellationAndRejectsSelectedFrameRequests() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecodeFrames(CreateTwoFrameGif(),
            new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token }, out _));
        Assert.Throws<ArgumentException>(() => OfficeRasterImageDecoder.TryDecodeFrames(CreateTwoFrameGif(),
            new OfficeRasterDecodeOptions { FrameIndex = 1 }, out _));
    }

    [Fact]
    public void FrameTransformPreservesTimingAndPlaybackWithSeparatelyOwnedImages() {
        var source = new OfficeRasterFrames(new[] {
            new OfficeRasterFrame(new OfficeRasterImage(2, 1, OfficeColor.Red), TimeSpan.FromMilliseconds(135)),
            new OfficeRasterFrame(new OfficeRasterImage(2, 1, OfficeColor.Lime), TimeSpan.FromMilliseconds(275))
        }, playCount: 3);

        OfficeRasterFrames result = source.Transform(image => OfficeRasterTransforms.Rotate(image, 90),
            image => (image.Height, image.Width));

        Assert.Equal(3, result.PlayCount);
        Assert.Equal(source.Select(frame => frame.Duration), result.Select(frame => frame.Duration));
        Assert.All(result, frame => { Assert.Equal(1, frame.Image.Width); Assert.Equal(2, frame.Image.Height); });
        result[0].Image.SetPixel(0, 0, OfficeColor.Blue);
        Assert.Equal(OfficeColor.Red, source[0].Image.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Lime, result[1].Image.GetPixel(0, 0));
    }

    [Fact]
    public void FrameTransformChecksTheEntirePlannedSequenceBeforeMapping() {
        var source = new OfficeRasterFrames(new[] {
            new OfficeRasterFrame(new OfficeRasterImage(1, 1)),
            new OfficeRasterFrame(new OfficeRasterImage(1, 1))
        });
        int calls = 0;
        Assert.Throws<ArgumentException>(() => source.Transform(image => { calls++; return image; },
            image => (5000, 6000)));
        Assert.Equal(0, calls);
        Assert.Throws<ArgumentException>(() => source.Transform(image => new OfficeRasterImage(2, 1)));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => source.Transform(image => { calls++; return image; },
            cancellationToken: cancellation.Token));
        Assert.Equal(0, calls);
    }

    [Fact]
    public void CanceledMultiPageEncodingLeavesTheDestinationOpenAndUnchanged() {
        var pages = new[] { new OfficeRasterImage(2, 1, OfficeColor.Red), new OfficeRasterImage(1, 2, OfficeColor.Blue) };
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var destination = new System.IO.MemoryStream();
        destination.WriteByte(123);
        Assert.Throws<OperationCanceledException>(() => OfficeTiffCodec.EncodePages(pages, null, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => OfficeTiffCodec.EncodePagesTo(pages, destination, null, cancellation.Token));
        Assert.Equal(new byte[] { 123 }, destination.ToArray());
        Assert.True(destination.CanWrite);
    }
}
