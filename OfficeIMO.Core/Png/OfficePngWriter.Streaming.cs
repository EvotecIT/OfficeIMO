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
            checkpointObserver);
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
        MemoryStream? ownedOutput = null) {
        cancellationToken.ThrowIfCancellationRequested();
        ValidateRgba(width, height, rgba);
        OfficeRasterOutput.EnsureWritable(destination);
        if (compression != OfficePngCompression.Optimal && compression != OfficePngCompression.Stored) {
            throw new ArgumentOutOfRangeException(nameof(compression));
        }

        bool bilevel = compression == OfficePngCompression.Optimal
            && IsOpaqueBilevel(rgba, cancellationToken, checkpointObserver);
        destination.Write(PngSignature, 0, PngSignature.Length);
        WriteChunk(destination, "IHDR", BuildIhdr(width, height, bilevel ? 1 : 8, bilevel ? 0 : 6));
        if (dpiX.HasValue && dpiY.HasValue) {
            WriteChunk(destination, "pHYs", BuildPhysicalResolution(dpiX.Value, dpiY.Value));
        }

        var idat = new PngIdatChunkStream(destination, StreamingIdatChunkSize);
        if (compression == OfficePngCompression.Optimal) {
            if (ownedOutput != null && !bilevel && rgba.Length >= MaterializedProbeMinimumRgbaBytes) {
                WriteMaterializedOptimalZlib(ownedOutput, idat, width, height, rgba,
                    cancellationToken, checkpointObserver);
            } else {
                WriteOptimalZlib(idat, width, height, rgba, bilevel, cancellationToken, checkpointObserver);
            }
        } else {
            WriteStoredZlib(idat, width, height, rgba, cancellationToken, checkpointObserver);
        }
        cancellationToken.ThrowIfCancellationRequested();
        idat.Complete();
        WriteChunk(destination, "IEND", Array.Empty<byte>());
    }

    private static void WriteOptimalZlib(
        PngIdatChunkStream destination,
        int width,
        int height,
        byte[] rgba,
        bool bilevel,
        System.Threading.CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        var workspace = new PngFilteringWorkspace(width, bilevel);
        using var adaptiveSize = new PngSizeProbeStream();
        using var unfilteredSize = new PngSizeProbeStream(destination);
        WriteRgbaZlib(adaptiveSize, height, rgba, workspace, adaptiveFiltering: true, cancellationToken, checkpointObserver);
        WriteRgbaZlib(unfilteredSize, height, rgba, workspace, adaptiveFiltering: false, cancellationToken, checkpointObserver);
        bool adaptiveFiltering = adaptiveSize.Length <= unfilteredSize.Length;
        cancellationToken.ThrowIfCancellationRequested();
        if (!adaptiveFiltering && unfilteredSize.FullyCaptured) return;
        destination.DiscardProbe();
        WriteRgbaZlib(destination, height, rgba, workspace,
            adaptiveFiltering, cancellationToken, checkpointObserver);
    }

    private static void WriteRgbaZlib(
        Stream destination, int height, byte[] rgba, PngFilteringWorkspace workspace,
        bool adaptiveFiltering, System.Threading.CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        destination.WriteByte(0x78);
        destination.WriteByte(0x9C);

        int stride = workspace.Stride;
        byte[] filteredRow = workspace.Row;
        byte[] paethCandidate = workspace.Paeth;
        byte[] compressionBatch = workspace.Batch;
        int batchLength = 0;
        uint adlerA = 1;
        uint adlerB = 0;

        using (var deflate = new DeflateStream(destination, CompressionLevel.Optimal, leaveOpen: true)) {
            for (int y = 0; y < height; y++) {
                checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngCompressionRow);
                cancellationToken.ThrowIfCancellationRequested();
                int rowOffset = y * stride;
                if (workspace.BilevelRows != null) {
                    FilterBilevelRow(rgba, y * workspace.RgbaStride, workspace,
                        y, adaptiveFiltering, cancellationToken, checkpointObserver);
                } else if (!adaptiveFiltering) {
                    filteredRow[0] = 0;
                    Buffer.BlockCopy(rgba, rowOffset, filteredRow, 1, stride);
                } else if (y == 0) {
                    filteredRow[0] = 1;
                    FilterFirstRowSub(rgba, rowOffset, stride, filteredRow, 1, cancellationToken, checkpointObserver);
                } else {
                    int previousRowOffset = rowOffset - stride;
                    long upScore = FilterUp(rgba, rowOffset, previousRowOffset, stride, filteredRow, 1, cancellationToken, checkpointObserver);
                    long paethScore = FilterPaeth(rgba, rowOffset, previousRowOffset, stride, paethCandidate, cancellationToken, checkpointObserver);
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
                UpdateAdler32(filteredRow, 0, filteredRow.Length, ref adlerA, ref adlerB, cancellationToken, checkpointObserver);
            }
            if (batchLength > 0) deflate.Write(compressionBatch, 0, batchLength);
        }

        WriteAdler32(destination, (adlerB << 16) | adlerA);
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

        // A size probe may retain its candidate in this existing bounded buffer.
        // It cannot write IDAT bytes to the caller before selection is complete.
        internal bool TryCaptureProbe(byte[] buffer, int offset, int count) {
            if (count > _chunkSize - _count) return false;
            EnsureCapacity(_count + count);
            Buffer.BlockCopy(buffer, offset, _buffer, _count, count);
            _count += count;
            return true;
        }

        internal bool TryCaptureProbeByte(byte value) {
            if (_count == _chunkSize) return false;
            EnsureCapacity(_count + 1);
            _buffer[_count++] = value;
            return true;
        }

#if NET8_0_OR_GREATER
        internal bool TryCaptureProbe(ReadOnlySpan<byte> buffer) {
            if (buffer.Length > _chunkSize - _count) return false;
            EnsureCapacity(_count + buffer.Length);
            buffer.CopyTo(_buffer.AsSpan(_count));
            _count += buffer.Length;
            return true;
        }
#endif

        internal void DiscardProbe() => _count = 0;

        // Only an encoder-owned materialized output can rewind a completed candidate.
        // Keep the bounded unfiltered probe so it can become the selected IDAT payload.
        internal void ResumeWriting() {
            if (!_completed) throw new InvalidOperationException("The PNG IDAT candidate is not complete.");
            _completed = false;
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
