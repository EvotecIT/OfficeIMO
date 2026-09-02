using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterStreamingEncodingTests {
    [Fact]
    public void MaterializedEncoderCallsWithPositionalNullRemainUnambiguous() {
        OfficeRasterImage image = CreateSampleImage();

        Assert.NotEmpty(OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png, null));
        Assert.Throws<ArgumentNullException>(() => OfficePngWriter.Encode(image, null!));
        Assert.NotEmpty(OfficeJpegCodec.Encode(image, null));
        Assert.NotEmpty(OfficeTiffCodec.Encode(image, null));
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void SharedEncoderWritesDeterministicOutputToForwardOnlyStreams(
        OfficeImageExportFormat format) {
        OfficeRasterImage image = CreateSampleImage();
        OfficeRasterEncodingOptions options = CreateOptions();
        byte[] first = EncodeToForwardOnlyStream(image, format, options);
        byte[] second = EncodeToForwardOnlyStream(image, format, options);

        Assert.Equal(first, second);
        OfficeImageInfo info = OfficeImageReader.Identify(first);
        Assert.Equal(ToImageFormat(format), info.Format);
        Assert.Equal(image.Width, info.Width);
        Assert.Equal(image.Height, info.Height);
        Assert.InRange(info.DpiX, 143.98D, 144.02D);
        Assert.InRange(info.DpiY, 119.98D, 120.02D);

        byte[] expected = OfficeRasterImageEncoder.Encode(image, format, options);
        AssertEquivalentPixels(expected, first);
    }

#if NET8_0_OR_GREATER
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void CodecBufferWriterOverloadsMatchCodecStreamOutput(
        OfficeImageExportFormat format) {
        OfficeRasterImage image = CreateSampleImage();
        OfficeRasterEncodingOptions options = CreateOptions();
        using var stream = new MemoryStream();
        EncodeWithCodec(image, format, stream, options);
        var writer = new CollectingBufferWriter();
        EncodeWithCodec(image, format, writer, options);

        Assert.Equal(stream.ToArray(), writer.ToArray());
    }
#endif

    [Theory]
    [InlineData(OfficeTiffCompression.None)]
    [InlineData(OfficeTiffCompression.Lzw)]
    [InlineData(OfficeTiffCompression.PackBits)]
    [InlineData(OfficeTiffCompression.Deflate)]
    public void TiffStreamEncodingMatchesMaterializedEncoding(
        OfficeTiffCompression compression) {
        OfficeRasterImage image = CreateSampleImage();
        var options = new OfficeTiffEncodeOptions {
            Compression = compression,
            DpiX = 144D,
            DpiY = 120D
        };
        byte[] expected = OfficeTiffCodec.Encode(image, options);
        using var actual = new MemoryStream();

        OfficeTiffCodec.EncodeTo(image, actual, options);

        Assert.Equal(expected, actual.ToArray());
    }

    [Theory]
    [InlineData(OfficePngCompression.Optimal)]
    [InlineData(OfficePngCompression.Stored)]
    public void PngStreamEncodingPreservesPixelsForEachCompression(
        OfficePngCompression compression) {
        OfficeRasterImage image = CreateSampleImage();
        using var actual = new MemoryStream();

        OfficePngWriter.EncodeTo(image, actual, compression);

        Assert.True(OfficePngReader.TryDecode(actual.ToArray(), out OfficeRasterImage? decoded));
        Assert.NotNull(decoded);
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
    }

    [Fact]
    public void SharedEncoderLeavesCallerStreamOpen() {
        OfficeRasterImage image = CreateSampleImage();
        using var destination = new MemoryStream();

        OfficeRasterImageEncoder.EncodeTo(
            image,
            OfficeImageExportFormat.Png,
            destination,
            CreateOptions());
        destination.WriteByte(0x7A);

        Assert.True(destination.Length > 1);
    }

    [Fact]
    public void SharedEncoderRejectsReadOnlyDestination() {
        OfficeRasterImage image = CreateSampleImage();
        using var destination = new MemoryStream(new byte[1], writable: false);

        Assert.Throws<ArgumentException>(() =>
            OfficeRasterImageEncoder.EncodeTo(
                image,
                OfficeImageExportFormat.Png,
                destination));
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void SharedEncoderEnforcesTheEncodedByteCeilingWhileWriting(
        OfficeImageExportFormat format) {
        OfficeImageExportBatchLimitException exception =
            Assert.Throws<OfficeImageExportBatchLimitException>(() =>
                OfficeRasterImageEncoder.Encode(
                    CreateSampleImage(),
                    format,
                    CreateOptions(),
                    maximumEncodedBytes: 8L,
                    cancellationToken: CancellationToken.None));

        Assert.Equal(nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes), exception.LimitName);
        Assert.True(exception.Actual > exception.Maximum);
        Assert.Equal(8L, exception.Maximum);
    }

    [Fact]
    public void MaterializedEncodingBudgetIncludesBackingBufferAndFinalCopy() {
        Assert.True(OfficeImageExportEncodingMemoryStream.IsFinalMaterializationWithinLimit(
            retainedManagedBytes: 8L * 1024L * 1024L,
            backingBytes: 16L * 1024L * 1024L,
            encodedLength: 12L * 1024L * 1024L));
        Assert.False(OfficeImageExportEncodingMemoryStream.IsFinalMaterializationWithinLimit(
            retainedManagedBytes: 128L * 1024L * 1024L,
            backingBytes: 128L * 1024L * 1024L,
            encodedLength: 128L * 1024L * 1024L));
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void SharedEncoderObservesCancellationBetweenWrites(OfficeImageExportFormat format) {
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancellingWriteStream(cancellation);

        Assert.Throws<OperationCanceledException>(() =>
            OfficeRasterImageEncoder.EncodeTo(
                CreateSampleImage(),
                format,
                destination,
                CreateOptions(),
                maximumEncodedBytes: long.MaxValue,
                cancellationToken: cancellation.Token));

        Assert.True(destination.WriteCount > 0);
    }

    [Fact]
    public void JpegCancellationCanStopCoefficientWorkBeforeTheFirstWrite() {
        OfficeRasterImage image = new OfficeRasterImage(2800, 2800, OfficeColor.CornflowerBlue);
        using var cancellation = new CancellationTokenSource();
        using var destination = new CountingWriteStream();
        cancellation.CancelAfter(1);

        Assert.Throws<OperationCanceledException>(() =>
            OfficeRasterImageEncoder.EncodeTo(
                image,
                OfficeImageExportFormat.Jpeg,
                destination,
                CreateOptions(),
                maximumEncodedBytes: long.MaxValue,
                cancellationToken: cancellation.Token));

        Assert.Equal(0, destination.WriteCount);
    }

    [Theory]
    [InlineData(OfficeTiffCompression.Lzw)]
    [InlineData(OfficeTiffCompression.PackBits)]
    [InlineData(OfficeTiffCompression.Deflate)]
    public void TiffCancellationCanStopCompressionBeforeTheFirstWrite(
        OfficeTiffCompression compression) {
        OfficeRasterImage image = new OfficeRasterImage(2800, 2800, OfficeColor.CornflowerBlue);
        using var cancellation = new CancellationTokenSource();
        using var destination = new CountingWriteStream();
        OfficeRasterEncodingOptions options = CreateOptions();
        options.Tiff.Compression = compression;
        cancellation.CancelAfter(1);

        Assert.Throws<OperationCanceledException>(() =>
            OfficeRasterImageEncoder.EncodeTo(
                image,
                OfficeImageExportFormat.Tiff,
                destination,
                options,
                maximumEncodedBytes: long.MaxValue,
                cancellationToken: cancellation.Token));

        Assert.Equal(0, destination.WriteCount);
    }

    private static byte[] EncodeToForwardOnlyStream(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        OfficeRasterEncodingOptions options) {
        using var destination = new ForwardOnlyWriteStream();
        OfficeRasterImageEncoder.EncodeTo(image, format, destination, options);
        destination.WriteByte(0x7A);
        byte[] withSentinel = destination.ToArray();
        Array.Resize(ref withSentinel, withSentinel.Length - 1);
        return withSentinel;
    }

#if NET8_0_OR_GREATER
    private static void EncodeWithCodec(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        Stream destination,
        OfficeRasterEncodingOptions options) {
        switch (format) {
            case OfficeImageExportFormat.Png:
                OfficePngWriter.EncodeTo(image, destination, options.Png);
                break;
            case OfficeImageExportFormat.Jpeg:
                OfficeJpegCodec.EncodeTo(image, destination, options.Jpeg);
                break;
            case OfficeImageExportFormat.Tiff:
                OfficeTiffCodec.EncodeTo(image, destination, options.Tiff);
                break;
            case OfficeImageExportFormat.Webp:
                OfficeWebpCodec.EncodeTo(image, destination, options.DpiX, options.DpiY);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }
    }
    private static void EncodeWithCodec(
        OfficeRasterImage image,
        OfficeImageExportFormat format,
        IBufferWriter<byte> destination,
        OfficeRasterEncodingOptions options) {
        switch (format) {
            case OfficeImageExportFormat.Png:
                OfficePngWriter.EncodeTo(image, destination, options.Png);
                break;
            case OfficeImageExportFormat.Jpeg:
                OfficeJpegCodec.EncodeTo(image, destination, options.Jpeg);
                break;
            case OfficeImageExportFormat.Tiff:
                OfficeTiffCodec.EncodeTo(image, destination, options.Tiff);
                break;
            case OfficeImageExportFormat.Webp:
                OfficeWebpCodec.EncodeTo(image, destination, options.DpiX, options.DpiY);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }
    }
#endif

    private static void AssertEquivalentPixels(byte[] expected, byte[] actual) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(expected, out OfficeRasterImage? expectedImage));
        Assert.True(OfficeRasterImageDecoder.TryDecode(actual, out OfficeRasterImage? actualImage));
        Assert.NotNull(expectedImage);
        Assert.NotNull(actualImage);
        Assert.Equal(expectedImage!.GetPixels(), actualImage!.GetPixels());
    }

    private static OfficeImageFormat ToImageFormat(OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => OfficeImageFormat.Png,
        OfficeImageExportFormat.Jpeg => OfficeImageFormat.Jpeg,
        OfficeImageExportFormat.Tiff => OfficeImageFormat.Tiff,
        OfficeImageExportFormat.Webp => OfficeImageFormat.Webp,
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };

    private static OfficeRasterEncodingOptions CreateOptions() => new() {
        DpiX = 144D,
        DpiY = 120D,
        Png = new OfficePngEncodeOptions {
            Compression = OfficePngCompression.Optimal
        },
        Jpeg = new OfficeJpegEncodeOptions {
            Quality = 85,
            Subsampling = OfficeJpegSubsampling.Y420,
            Progressive = true,
            OptimizeHuffman = true,
            Background = OfficeColor.White
        },
        Tiff = new OfficeTiffEncodeOptions {
            Compression = OfficeTiffCompression.PackBits
        }
    };

    private static OfficeRasterImage CreateSampleImage() {
        var image = new OfficeRasterImage(19, 13);
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                image.SetPixel(
                    x,
                    y,
                    OfficeColor.FromRgba(
                        (byte)(x * 11),
                        (byte)(y * 17),
                        (byte)((x * 7 + y * 13) & 255),
                        (byte)(64 + ((x * 9 + y * 5) & 191))));
            }
        }
        return image;
    }

    private sealed class ForwardOnlyWriteStream : Stream {
        private readonly MemoryStream _inner = new();

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => _inner.Length;

        public override long Position {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        internal byte[] ToArray() => _inner.ToArray();
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
        public override void WriteByte(byte value) => _inner.WriteByte(value);

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class CancellingWriteStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;

        internal CancellingWriteStream(CancellationTokenSource cancellation) {
            _cancellation = cancellation;
        }

        internal int WriteCount { get; private set; }

        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            WriteCount++;
            _cancellation.Cancel();
        }

#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) {
            base.Write(buffer);
            WriteCount++;
            _cancellation.Cancel();
        }
#endif

        public override void WriteByte(byte value) {
            base.WriteByte(value);
            WriteCount++;
            _cancellation.Cancel();
        }
    }

    private sealed class CountingWriteStream : MemoryStream {
        internal int WriteCount { get; private set; }

        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            WriteCount++;
        }

#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) {
            base.Write(buffer);
            WriteCount++;
        }
#endif

        public override void WriteByte(byte value) {
            base.WriteByte(value);
            WriteCount++;
        }
    }

#if NET8_0_OR_GREATER
    private sealed class CollectingBufferWriter : IBufferWriter<byte> {
        private readonly MemoryStream _output = new();
        private byte[] _buffer = new byte[256];

        public void Advance(int count) {
            if (count < 0 || count > _buffer.Length) throw new ArgumentOutOfRangeException(nameof(count));
            _output.Write(_buffer, 0, count);
        }

        public Memory<byte> GetMemory(int sizeHint = 0) {
            EnsureCapacity(sizeHint);
            return new Memory<byte>(_buffer);
        }

        public Span<byte> GetSpan(int sizeHint = 0) {
            EnsureCapacity(sizeHint);
            return new Span<byte>(_buffer);
        }

        internal byte[] ToArray() => _output.ToArray();

        private void EnsureCapacity(int sizeHint) {
            if (sizeHint < 0) throw new ArgumentOutOfRangeException(nameof(sizeHint));
            if (sizeHint > _buffer.Length) _buffer = new byte[sizeHint];
        }
    }
#endif
}
