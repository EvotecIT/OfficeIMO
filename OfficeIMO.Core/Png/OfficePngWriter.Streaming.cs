using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;
using System.IO.Compression;

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    private const int StreamingIdatChunkSize = 64 * 1024;

    /// <summary>Encodes an RGBA image directly to a caller-owned writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        OfficePngCompression compression = OfficePngCompression.Optimal) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        EncodeRgbaStreaming(
            image.Width,
            image.Height,
            image.PixelBuffer,
            destination,
            compression,
            dpiX: null,
            dpiY: null,
            System.Threading.CancellationToken.None);
    }

    /// <summary>Encodes an RGBA image directly to a writable stream with cooperative cancellation.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        System.Threading.CancellationToken cancellationToken,
        OfficePngCompression compression = OfficePngCompression.Optimal) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        EncodeRgbaStreaming(
            image.Width,
            image.Height,
            image.PixelBuffer,
            destination,
            compression,
            dpiX: null,
            dpiY: null,
            cancellationToken);
    }

    /// <summary>Encodes an RGBA image with physical-resolution metadata directly to a writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        OfficePngEncodeOptions options) {
        EncodeTo(image, destination, options, System.Threading.CancellationToken.None);
    }

    /// <summary>Encodes an RGBA image with physical-resolution metadata and cooperative cancellation.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        OfficePngEncodeOptions options,
        System.Threading.CancellationToken cancellationToken) {
        EncodeTo(image, destination, options, cancellationToken, checkpointObserver: null);
    }

    internal static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        OfficePngEncodeOptions options,
        System.Threading.CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        if (options == null) throw new ArgumentNullException(nameof(options));
        ValidateDpi(options.DpiX, nameof(options.DpiX));
        ValidateDpi(options.DpiY, nameof(options.DpiY));
        EncodeRgbaStreaming(
            image.Width,
            image.Height,
            image.PixelBuffer,
            destination,
            options.Compression,
            options.WritePhysicalResolution ? options.DpiX : (double?)null,
            options.WritePhysicalResolution ? options.DpiY : (double?)null,
            cancellationToken,
            checkpointObserver, requireRgbColorSpace: options.RequireRgbColorSpace, requireGrayColorSpace: options.RequireGrayColorSpace, effort: options.Effort);
    }

#if NET8_0_OR_GREATER
    /// <summary>Encodes an RGBA image directly to a caller-owned buffer writer.</summary>
    public static void EncodeTo(
        OfficeRasterImage image,
        IBufferWriter<byte> destination,
        OfficePngCompression compression = OfficePngCompression.Optimal) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        using var stream = new OfficeBufferWriterStream(destination);
        EncodeTo(image, stream, compression);
    }

    /// <summary>Encodes an RGBA image with physical-resolution metadata directly to a buffer writer.</summary>
    public static void EncodeTo(
        OfficeRasterImage image,
        IBufferWriter<byte> destination,
        OfficePngEncodeOptions options) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        using var stream = new OfficeBufferWriterStream(destination);
        EncodeTo(image, stream, options);
    }
#endif

    private static void EncodeRgbaStreaming(
        int width,
        int height,
        byte[] rgba,
        Stream destination,
        OfficePngCompression compression,
        double? dpiX,
        double? dpiY,
        System.Threading.CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null,
        bool requireRgbColorSpace = false, bool requireGrayColorSpace = false,
        OfficeImageEncodingEffort effort = OfficeImageEncodingEffort.Balanced) {
        cancellationToken.ThrowIfCancellationRequested();
        ValidateRgba(width, height, rgba);
        ValidatePngEffort(effort);
        OfficeRasterOutput.EnsureWritable(destination);
        if (compression != OfficePngCompression.Optimal && compression != OfficePngCompression.Stored && compression != OfficePngCompression.Fastest) {
            throw new ArgumentOutOfRangeException(nameof(compression));
        }

        if (compression == OfficePngCompression.Optimal || requireGrayColorSpace) {
            EncodeAdaptiveRgba(width, height, rgba, destination, dpiX, dpiY,
                cancellationToken, checkpointObserver, requireRgbColorSpace, requireGrayColorSpace, compression, effort);
            return;
        }
        destination.Write(PngSignature, 0, PngSignature.Length);
        WriteChunk(destination, "IHDR", BuildIhdr(width, height, 8, 6));
        WriteSrgbChunks(destination);
        if (dpiX.HasValue && dpiY.HasValue)
            WriteChunk(destination, "pHYs", BuildPhysicalResolution(dpiX.Value, dpiY.Value));
        var idat = new PngIdatChunkStream(destination, StreamingIdatChunkSize);
        if (compression == OfficePngCompression.Fastest) {
            var layout = new PngColorLayout(6, 8);
            WriteColorZlib(idat, width, height, rgba, null, layout, false,
                cancellationToken, checkpointObserver, CompressionLevel.Fastest);
        } else {
            WriteStoredZlib(idat, width, height, rgba, cancellationToken, checkpointObserver);
        }
        cancellationToken.ThrowIfCancellationRequested();
        idat.Complete();
        WriteChunk(destination, "IEND", Array.Empty<byte>());
    }

    private static void WriteStoredZlib(
        Stream destination,
        int width,
        int height,
        byte[] rgba,
        System.Threading.CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        destination.WriteByte(0x78);
        destination.WriteByte(0x01);

        int stride = checked(width * 4);
        int totalLength = checked(height * (stride + 1));
        var block = new byte[Math.Min(65535, totalLength)];
        int row = 0;
        int rowPosition = -1;
        int remaining = totalLength;
        uint adlerA = 1;
        uint adlerB = 0;

        while (remaining > 0) {
            checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngCompressionRow);
            cancellationToken.ThrowIfCancellationRequested();
            int blockLength = Math.Min(65535, remaining);
            int target = 0;
            while (target < blockLength) {
                if (rowPosition < 0) {
                    block[target++] = 0;
                    rowPosition = 0;
                    continue;
                }

                int take = Math.Min(stride - rowPosition, blockLength - target);
                Buffer.BlockCopy(rgba, checked(row * stride + rowPosition), block, target, take);
                rowPosition += take;
                target += take;
                if (rowPosition == stride) {
                    row++;
                    rowPosition = -1;
                }
            }

            remaining -= blockLength;
            destination.WriteByte(remaining == 0 ? (byte)1 : (byte)0);
            destination.WriteByte((byte)blockLength);
            destination.WriteByte((byte)(blockLength >> 8));
            ushort inverse = unchecked((ushort)~blockLength);
            destination.WriteByte((byte)inverse);
            destination.WriteByte((byte)(inverse >> 8));
            destination.Write(block, 0, blockLength);
            UpdateAdler32(block, 0, blockLength, ref adlerA, ref adlerB, cancellationToken, checkpointObserver);
        }

        WriteAdler32(destination, (adlerB << 16) | adlerA);
    }

    private static void WriteAdler32(Stream destination, uint adler) {
        destination.WriteByte((byte)(adler >> 24));
        destination.WriteByte((byte)(adler >> 16));
        destination.WriteByte((byte)(adler >> 8));
        destination.WriteByte((byte)adler);
    }

    private sealed class PngIdatChunkStream : Stream {
        private readonly Stream _destination;
        private readonly int _chunkSize;
        private byte[] _buffer;
        private int _count;
        private bool _completed;

        internal PngIdatChunkStream(Stream destination, int chunkSize) {
            _destination = destination;
            _chunkSize = chunkSize;
            // Grow with the compressed output, including large rasters that compress to a tiny PNG.
            _buffer = new byte[Math.Min(256, chunkSize)];
        }

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => !_completed;
        public override long Length => throw new NotSupportedException();

        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        internal void Complete() {
            if (_completed) return;
            FlushChunk();
            _completed = true;
        }

        public override void Flush() => _destination.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) {
            if (_completed) throw new InvalidOperationException("The PNG IDAT stream is complete.");
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0) throw new ArgumentOutOfRangeException(nameof(offset));
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            if (offset > buffer.Length - count) throw new ArgumentException("The buffer range is invalid.", nameof(buffer));

            while (count > 0) {
                int copied = Math.Min(count, _chunkSize - _count);
                EnsureCapacity(_count + copied);
                Buffer.BlockCopy(buffer, offset, _buffer, _count, copied);
                _count += copied;
                offset += copied;
                count -= copied;
                if (_count == _chunkSize) FlushChunk();
            }
        }

        public override void WriteByte(byte value) {
            if (_completed) throw new InvalidOperationException("The PNG IDAT stream is complete.");
            EnsureCapacity(_count + 1);
            _buffer[_count++] = value;
            if (_count == _chunkSize) FlushChunk();
        }

        private void EnsureCapacity(int required) {
            if (required <= _buffer.Length) return;
            int capacity = Math.Min(_chunkSize, Math.Max(required, _buffer.Length * 2));
            var buffer = new byte[capacity];
            Buffer.BlockCopy(_buffer, 0, buffer, 0, _count);
            _buffer = buffer;
        }

        private void FlushChunk() {
            if (_count == 0) return;
            if (_count == _buffer.Length) {
                WriteChunk(_destination, "IDAT", _buffer);
            } else {
                var chunk = new byte[_count];
                Buffer.BlockCopy(_buffer, 0, chunk, 0, _count);
                WriteChunk(_destination, "IDAT", chunk);
            }
            _count = 0;
        }
    }
}
