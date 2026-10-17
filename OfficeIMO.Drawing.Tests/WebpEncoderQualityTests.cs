using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WebpEncoderQualityTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UncorrelatedRgbaIncludingFullyTransparentNoiseStaysExactAndBounded(bool transparent) {
        const int width = 129, height = 65;
        var pixels = new byte[width * height * 4];
        uint random = 0x39BE713DU;
        for (int offset = 0; offset < pixels.Length; offset += 4) {
            uint color = Next(ref random);
            pixels[offset] = (byte)color;
            pixels[offset + 1] = (byte)(color >> 8);
            pixels[offset + 2] = (byte)(color >> 16);
            pixels[offset + 3] = transparent ? (byte)0 : (byte)(color >> 24);
        }
        var image = OfficeRasterImage.FromRgba32(width, height, pixels);
        using var streamed = new MemoryStream();
        OfficeWebpCodec.EncodeTo(image, streamed);

        byte[] encoded = OfficeWebpCodec.Encode(image);

        Assert.True(encoded.LongLength <= streamed.Length);
        AssertExact(image, encoded);
        AssertExact(image, streamed.ToArray());
        Assert.Equal(pixels, image.GetPixels());
    }

    [Fact]
    public void UnorderedRepeatedColorsCompressWhilePreservingHiddenRgb() {
        const int width = 193, height = 97;
        var pixels = new byte[width * height * 4];
        uint random = 0x491A36FBU;
        for (int pixel = 0; pixel < width * height; pixel++) {
            int color = (int)(Next(ref random) & 31);
            pixels[pixel * 4] = (byte)(color * 73);
            pixels[pixel * 4 + 1] = (byte)(color * 151);
            pixels[pixel * 4 + 2] = (byte)(color * 199);
            pixels[pixel * 4 + 3] = (byte)((color & 3) == 0 ? 0 : color * 7);
        }
        var image = OfficeRasterImage.FromRgba32(width, height, pixels);

        byte[] encoded = OfficeWebpCodec.Encode(image);

        Assert.True(encoded.Length < pixels.Length / 3, $"Expected palette reuse, got {encoded.Length} bytes for {pixels.Length} RGBA bytes.");
        AssertExact(image, encoded);
        Assert.Equal(encoded, OfficeWebpCodec.Encode(image));
        Assert.Equal(pixels, image.GetPixels());
    }

    [Theory]
    [InlineData(1, 129)]
    [InlineData(129, 1)]
    [InlineData(33, 65)]
    [InlineData(97, 67)]
    public void BlockEdgesAndThinImagesPreserveEveryRgbaSample(int width, int height) {
        var pixels = new byte[width * height * 4];
        uint random = 0x812357E1U;
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                int offset = (y * width + x) * 4;
                uint noise = Next(ref random);
                int region = ((x / 32) + (y / 32)) % 3;
                pixels[offset] = region == 0 ? (byte)(x * 3 + y) : (byte)noise;
                pixels[offset + 1] = region == 1 ? (byte)(x + y * 5) : (byte)(noise >> 8);
                pixels[offset + 2] = region == 2 ? (byte)(x * 7 - y * 3) : (byte)(noise >> 16);
                pixels[offset + 3] = (x + y) % 5 == 0 ? (byte)0 : (byte)(noise >> 24);
            }
        }
        var image = OfficeRasterImage.FromRgba32(width, height, pixels);

        AssertExact(image, OfficeWebpCodec.Encode(image));
    }

    [Fact]
    public void RepeatedTilesSurviveLongCopiesAndMatchWindowRollover() {
        const int width = 257, height = 300, tilePixels = 509;
        var pixels = new byte[width * height * 4];
        uint random = 0xCA39175BU;
        for (int pixel = 0; pixel < tilePixels; pixel++) {
            uint color = Next(ref random);
            for (int channel = 0; channel < 4; channel++) pixels[pixel * 4 + channel] = (byte)(color >> (channel * 8));
        }
        for (int pixel = tilePixels; pixel < width * height; pixel++) {
            Buffer.BlockCopy(pixels, (pixel % tilePixels) * 4, pixels, pixel * 4, 4);
        }
        var image = OfficeRasterImage.FromRgba32(width, height, pixels);

        byte[] encoded = OfficeWebpCodec.Encode(image);

        Assert.True(encoded.Length < pixels.Length / 8);
        AssertExact(image, encoded);
    }

    [Fact]
    public void CorrelatedSignedChannelsRoundTripAcrossColorTransformWraparound() {
        const int width = 131, height = 99;
        var pixels = new byte[width * height * 4];
        uint random = 0x1839ACE7U;
        for (int pixel = 0; pixel < width * height; pixel++) {
            uint value = Next(ref random);
            int green = (int)value & 255;
            int red = (unchecked((sbyte)green) * 3 >> 1) + (int)((value >> 8) & 3);
            int blue = (unchecked((sbyte)green) >> 1) + unchecked((sbyte)red);
            pixels[pixel * 4] = (byte)red;
            pixels[pixel * 4 + 1] = (byte)green;
            pixels[pixel * 4 + 2] = (byte)blue;
            pixels[pixel * 4 + 3] = (pixel & 3) == 0 ? (byte)0 : (byte)(value >> 24);
        }
        var image = OfficeRasterImage.FromRgba32(width, height, pixels);

        AssertExact(image, OfficeWebpCodec.Encode(image));
    }

    [Fact]
    public void CompressionWorkspaceFallsBackBeforeExceedingRetainedMemoryBudget() {
        var image = new OfficeRasterImage(64, 64, OfficeColor.FromRgba(211, 73, 19, 0));
        var options = new OfficeWebpEncodeOptions {
            RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes - 256 * 1024L
        };
        using var literal = new MemoryStream();
        OfficeWebpCodec.EncodeTo(image, literal, options);

        byte[] encoded = OfficeWebpCodec.Encode(image, options);

        Assert.Equal(literal.ToArray(), encoded);
        Assert.True(encoded.Length > image.Width * image.Height * 4);
        AssertExact(image, encoded);
    }

    [Fact]
    public void ForwardOnlyStreamsUseTheCompressedCandidateAndLeaveTheDestinationOpen() {
        var image = new OfficeRasterImage(257, 131, OfficeColor.FromRgba(93, 174, 211, 0));
        using var destination = new ForwardOnlyWebpStream();

        OfficeWebpCodec.EncodeTo(image, destination);

        byte[] encoded = destination.ToArray();
        Assert.True(destination.CanWrite);
        Assert.InRange(destination.MaximumWrite, 1, 16 * 1024);
        Assert.True(encoded.Length < image.Width * image.Height / 8);
        Assert.Equal(OfficeWebpCodec.Encode(image), encoded);
        AssertExact(image, encoded);
#if NET8_0_OR_GREATER
        var buffer = new ArrayBufferWriter<byte>();
        OfficeWebpCodec.EncodeTo(image, buffer);
        Assert.Equal(encoded, buffer.WrittenSpan.ToArray());
#endif
    }

    [Fact]
    public void CompressedStreamStopsWhenTheDestinationCancelsAfterWritingAPayloadChunk() {
        var image = new OfficeRasterImage(128, 128, OfficeColor.CornflowerBlue);
        using var cancellation = new CancellationTokenSource();
        using var destination = new ForwardOnlyWebpStream(cancellation);

        var error = Assert.Throws<OperationCanceledException>(() => OfficeWebpCodec.EncodeTo(image, destination,
            new OfficeWebpEncodeOptions(), cancellation.Token));

        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal(2, destination.WriteCount); // RIFF header and compressed payload
        Assert.True(destination.CanWrite);
    }

    [Fact]
    public void CompressedStreamObservesEncodedByteLimitsDuringChunkWrites() {
        var image = new OfficeRasterImage(128, 128, OfficeColor.FromRgba(211, 61, 32, 0));
        using var destination = new ForwardOnlyWebpStream();

        Assert.Throws<OfficeImageExportBatchLimitException>(() => OfficeRasterImageEncoder.EncodeTo(image,
            OfficeImageExportFormat.Webp, destination, new OfficeRasterEncodingOptions(), maximumEncodedBytes: 20));

        Assert.InRange(destination.WriteCount, 0, 1);
        Assert.True(destination.ToArray().Length <= 20);
        Assert.True(destination.CanWrite);
    }

    private sealed class ForwardOnlyWebpStream : Stream {
        private readonly MemoryStream _bytes = new MemoryStream();
        private readonly CancellationTokenSource? _cancellation;
        private bool _disposed;

        internal ForwardOnlyWebpStream(CancellationTokenSource? cancellation = null) => _cancellation = cancellation;
        internal int MaximumWrite { get; private set; }
        internal int WriteCount { get; private set; }
        internal byte[] ToArray() => _bytes.ToArray();
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => !_disposed;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) {
            if (_disposed) throw new ObjectDisposedException(nameof(ForwardOnlyWebpStream));
            _bytes.Write(buffer, offset, count);
            MaximumWrite = Math.Max(MaximumWrite, count);
            if (++WriteCount == 2) _cancellation?.Cancel();
        }
        protected override void Dispose(bool disposing) {
            if (disposing) { _disposed = true; _bytes.Dispose(); }
            base.Dispose(disposing);
        }
    }

    private static uint Next(ref uint state) {
        state ^= state << 13;
        state ^= state >> 17;
        state ^= state << 5;
        return state;
    }

    private static void AssertExact(OfficeRasterImage expected, byte[] encoded) {
        Assert.True(OfficeWebpCodec.TryDecode(encoded, out OfficeRasterImage? actual));
        Assert.NotNull(actual);
        Assert.Equal((expected.Width, expected.Height), (actual!.Width, actual.Height));
        Assert.Equal(expected.GetPixels(), actual.GetPixels());
    }
}
