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
    public static void Encode(
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
            dpiY: null);
    }

    /// <summary>Encodes an RGBA image with physical-resolution metadata directly to a writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void Encode(
        OfficeRasterImage image,
        Stream destination,
        OfficePngEncodeOptions options) {
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
            options.DpiX,
            options.DpiY);
    }

#if NET8_0_OR_GREATER
    /// <summary>Encodes an RGBA image directly to a caller-owned buffer writer.</summary>
    public static void Encode(
        OfficeRasterImage image,
        IBufferWriter<byte> destination,
        OfficePngCompression compression = OfficePngCompression.Optimal) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        using var stream = new OfficeBufferWriterStream(destination);
        Encode(image, stream, compression);
    }

    /// <summary>Encodes an RGBA image with physical-resolution metadata directly to a buffer writer.</summary>
    public static void Encode(
        OfficeRasterImage image,
        IBufferWriter<byte> destination,
        OfficePngEncodeOptions options) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        using var stream = new OfficeBufferWriterStream(destination);
        Encode(image, stream, options);
    }
#endif

    private static void EncodeRgbaStreaming(
        int width,
        int height,
        byte[] rgba,
        Stream destination,
        OfficePngCompression compression,
        double? dpiX,
        double? dpiY) {
        ValidateRgba(width, height, rgba);
        OfficeRasterOutput.EnsureWritable(destination);
        if (compression != OfficePngCompression.Optimal && compression != OfficePngCompression.Stored) {
            throw new ArgumentOutOfRangeException(nameof(compression));
        }

        destination.Write(PngSignature, 0, PngSignature.Length);
        WriteChunk(destination, "IHDR", BuildIhdr(width, height, 8, 6));
        if (dpiX.HasValue && dpiY.HasValue) {
            WriteChunk(destination, "pHYs", BuildPhysicalResolution(dpiX.Value, dpiY.Value));
        }

        var idat = new PngIdatChunkStream(destination, StreamingIdatChunkSize);
        if (compression == OfficePngCompression.Optimal) {
            WriteOptimalZlib(idat, width, height, rgba);
        } else {
            WriteStoredZlib(idat, width, height, rgba);
        }
        idat.Complete();
        WriteChunk(destination, "IEND", Array.Empty<byte>());
    }

    private static void WriteOptimalZlib(Stream destination, int width, int height, byte[] rgba) {
        destination.WriteByte(0x78);
        destination.WriteByte(0x9C);

        int stride = checked(width * 4);
        var filteredRow = new byte[checked(stride + 1)];
        var paethCandidate = new byte[stride];
        var compressionBatch = new byte[Math.Max(filteredRow.Length, 64 * 1024)];
        int batchLength = 0;
        uint adlerA = 1;
        uint adlerB = 0;

        using (var deflate = new DeflateStream(destination, CompressionLevel.Optimal, leaveOpen: true)) {
            for (int y = 0; y < height; y++) {
                int rowOffset = y * stride;
                if (y == 0) {
                    filteredRow[0] = 1;
                    FilterFirstRowSub(rgba, rowOffset, stride, filteredRow, 1);
                } else {
                    int previousRowOffset = rowOffset - stride;
                    long upScore = FilterUp(rgba, rowOffset, previousRowOffset, stride, filteredRow, 1);
                    long paethScore = FilterPaeth(rgba, rowOffset, previousRowOffset, stride, paethCandidate);
                    if (paethScore < upScore) {
                        filteredRow[0] = 4;
                        Buffer.BlockCopy(paethCandidate, 0, filteredRow, 1, stride);
                    } else {
                        filteredRow[0] = 2;
                    }
                }

                if (filteredRow.Length > compressionBatch.Length - batchLength) {
                    deflate.Write(compressionBatch, 0, batchLength);
                    batchLength = 0;
                }
                Buffer.BlockCopy(filteredRow, 0, compressionBatch, batchLength, filteredRow.Length);
                batchLength += filteredRow.Length;
                UpdateAdler32(filteredRow, 0, filteredRow.Length, ref adlerA, ref adlerB);
            }
            if (batchLength > 0) deflate.Write(compressionBatch, 0, batchLength);
        }

        WriteAdler32(destination, (adlerB << 16) | adlerA);
    }

    private static void WriteStoredZlib(Stream destination, int width, int height, byte[] rgba) {
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
            UpdateAdler32(block, 0, blockLength, ref adlerA, ref adlerB);
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
        private readonly byte[] _buffer;
        private int _count;
        private bool _completed;

        internal PngIdatChunkStream(Stream destination, int chunkSize) {
            _destination = destination;
            _buffer = new byte[chunkSize];
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
                int copied = Math.Min(count, _buffer.Length - _count);
                Buffer.BlockCopy(buffer, offset, _buffer, _count, copied);
                _count += copied;
                offset += copied;
                count -= copied;
                if (_count == _buffer.Length) FlushChunk();
            }
        }

        public override void WriteByte(byte value) {
            if (_completed) throw new InvalidOperationException("The PNG IDAT stream is complete.");
            _buffer[_count++] = value;
            if (_count == _buffer.Length) FlushChunk();
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
