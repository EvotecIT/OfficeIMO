using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    /// <summary>Encodes an RGBA image directly to a caller-owned writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(OfficeRasterImage image, Stream destination) {
        EncodeStreaming(image, destination, includeResolutionMetadata: false, 96D, 96D, CancellationToken.None);
    }

    /// <summary>Encodes an RGBA image with Exif resolution metadata directly to a writable stream.</summary>
    /// <remarks>The destination remains open after encoding.</remarks>
    public static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        double dpiX,
        double dpiY) {
        ValidateDpi(dpiX, nameof(dpiX));
        ValidateDpi(dpiY, nameof(dpiY));
        EncodeStreaming(image, destination, includeResolutionMetadata: true, dpiX, dpiY, CancellationToken.None);
    }

    internal static void EncodeTo(
        OfficeRasterImage image,
        Stream destination,
        double? dpiX,
        double? dpiY,
        CancellationToken cancellationToken) {
        bool writeResolution = dpiX.HasValue && dpiY.HasValue;
        if (writeResolution) {
            ValidateDpi(dpiX!.Value, nameof(dpiX));
            ValidateDpi(dpiY!.Value, nameof(dpiY));
        }
        EncodeStreaming(image, destination, writeResolution, dpiX ?? 96D, dpiY ?? 96D, cancellationToken);
    }

#if NET8_0_OR_GREATER
    /// <summary>Encodes an RGBA image directly to a caller-owned buffer writer.</summary>
    public static void EncodeTo(OfficeRasterImage image, IBufferWriter<byte> destination) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        using var stream = new OfficeBufferWriterStream(destination);
        EncodeTo(image, stream);
    }

    /// <summary>Encodes an RGBA image with Exif resolution metadata directly to a buffer writer.</summary>
    public static void EncodeTo(
        OfficeRasterImage image,
        IBufferWriter<byte> destination,
        double dpiX,
        double dpiY) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        using var stream = new OfficeBufferWriterStream(destination);
        EncodeTo(image, stream, dpiX, dpiY);
    }
#endif

    private static void EncodeStreaming(
        OfficeRasterImage image,
        Stream destination,
        bool includeResolutionMetadata,
        double dpiX,
        double dpiY,
        CancellationToken cancellationToken) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        OfficeRasterOutput.EnsureWritable(destination);
        cancellationToken.ThrowIfCancellationRequested();
        if (image.Width > OfficeRasterImageEncoder.WebpMaximumDimension) throw new ArgumentOutOfRangeException(nameof(image), "WebP width cannot exceed 16,384 pixels.");
        if (image.Height > OfficeRasterImageEncoder.WebpMaximumDimension) throw new ArgumentOutOfRangeException(nameof(image), "WebP height cannot exceed 16,384 pixels.");

        byte[] pixels = image.PixelBuffer;
        bool hasAlpha = HasTransparency(pixels, cancellationToken);
        int payloadLength = checked(1 + (int)((LiteralHeaderBitCount + pixels.LongLength * 8L + 7L) / 8L));
        int paddedPayloadLength = checked(payloadLength + (payloadLength & 1));
        byte[]? exif = includeResolutionMetadata
            ? CreateResolutionExif(image.Width, image.Height, dpiX, dpiY)
            : null;
        int paddedExifLength = exif == null ? 0 : checked(exif.Length + (exif.Length & 1));
        int fileLength = exif == null
            ? checked(20 + paddedPayloadLength)
            : checked(12 + 18 + 8 + paddedPayloadLength + 8 + paddedExifLength);
        if (fileLength > OfficeRasterGuards.MaximumEncodedBytes) {
            throw new ArgumentException("WebP output exceeds encoded-size limits.", nameof(image));
        }
        try {
            long retainedOutputCopies = OfficeRasterOutput.TryGetMemoryStream(destination, out _) ? 2L : 0L;
            long peakBytes = checked(
                pixels.LongLength + 24L +
                retainedOutputCopies * (fileLength + 24L) +
                (exif?.LongLength ?? 0L) + 24L +
                16L * 1024L);
            if (peakBytes > OfficeRasterGuards.MaximumDecodedBytes) {
                throw new ArgumentException("WebP encoding exceeds the managed working-set limit.", nameof(image));
            }
        } catch (OverflowException) {
            throw new ArgumentException("WebP encoding exceeds the managed working-set limit.", nameof(image));
        }

        byte[] header;
        if (exif == null) {
            header = new byte[20];
            WriteAscii(header, 0, "RIFF");
            WriteUInt32(header, 4, fileLength - 8);
            WriteAscii(header, 8, "WEBP");
            WriteAscii(header, 12, "VP8L");
            WriteUInt32(header, 16, payloadLength);
        } else {
            header = new byte[38];
            WriteAscii(header, 0, "RIFF");
            WriteUInt32(header, 4, fileLength - 8);
            WriteAscii(header, 8, "WEBP");
            WriteAscii(header, 12, "VP8X");
            WriteUInt32(header, 16, 10);
            header[20] = (byte)(0x08 | (hasAlpha ? 0x10 : 0x00));
            WriteUInt24(header, 24, image.Width - 1);
            WriteUInt24(header, 27, image.Height - 1);
            WriteAscii(header, 30, "VP8L");
            WriteUInt32(header, 34, payloadLength);
        }
        destination.Write(header, 0, header.Length);

        destination.WriteByte(0x2F);
        using var writer = new StreamLsbBitWriter(destination);
        writer.WriteBits((uint)(image.Width - 1), 14);
        writer.WriteBits((uint)(image.Height - 1), 14);
        writer.WriteBits(hasAlpha ? 1U : 0U, 1);
        writer.WriteBits(0, 3);
        writer.WriteBits(0, 1);
        writer.WriteBits(0, 1);
        writer.WriteBits(0, 1);
        WriteLiteralTree(writer, 280);
        WriteLiteralTree(writer, 256);
        WriteLiteralTree(writer, 256);
        WriteLiteralTree(writer, 256);
        WriteSingleSymbolTree(writer);
        for (int offset = 0; offset < pixels.Length; offset += 4) {
            if ((offset & 0x3FFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            writer.WriteBits(ReverseByte(pixels[offset + 1]), 8);
            writer.WriteBits(ReverseByte(pixels[offset]), 8);
            writer.WriteBits(ReverseByte(pixels[offset + 2]), 8);
            writer.WriteBits(ReverseByte(pixels[offset + 3]), 8);
        }
        writer.Flush();
        if (writer.BytesWritten != payloadLength - 1) {
            throw new InvalidOperationException("The literal WebP payload length calculation is inconsistent.");
        }
        if ((payloadLength & 1) != 0) destination.WriteByte(0);

        if (exif != null) {
            var exifHeader = new byte[8];
            WriteAscii(exifHeader, 0, "EXIF");
            WriteUInt32(exifHeader, 4, exif.Length);
            destination.Write(exifHeader, 0, exifHeader.Length);
            destination.Write(exif, 0, exif.Length);
            if ((exif.Length & 1) != 0) destination.WriteByte(0);
        }
    }

    private interface ILsbBitWriter {
        void WriteBits(uint value, int count);
        void Flush();
    }

    private sealed class StreamLsbBitWriter : ILsbBitWriter, IDisposable {
        private const int OutputBufferSize = 16 * 1024;
        private readonly Stream _destination;
        private byte[]? _output = new byte[OutputBufferSize];
        private int _outputCount;
        private ulong _buffer;
        private int _bitCount;

        internal StreamLsbBitWriter(Stream destination) {
            _destination = destination;
        }

        internal int BytesWritten { get; private set; }

        public void WriteBits(uint value, int count) {
            if (count < 0 || count > 32) throw new ArgumentOutOfRangeException(nameof(count));
            ulong mask = count == 32 ? uint.MaxValue : ((1UL << count) - 1UL);
            _buffer |= ((ulong)value & mask) << _bitCount;
            _bitCount += count;
            while (_bitCount >= 8) {
                WriteByte((byte)_buffer);
                _buffer >>= 8;
                _bitCount -= 8;
            }
        }

        public void Flush() {
            if (_bitCount > 0) {
                WriteByte((byte)_buffer);
                _buffer = 0;
                _bitCount = 0;
            }
            FlushOutput();
        }

        public void Dispose() {
            byte[]? output = _output;
            if (output == null) return;
            try {
                Flush();
            } finally {
                _output = null;
            }
        }

        private void WriteByte(byte value) {
            byte[] output = _output ?? throw new ObjectDisposedException(nameof(StreamLsbBitWriter));
            output[_outputCount++] = value;
            BytesWritten = checked(BytesWritten + 1);
            if (_outputCount == output.Length) FlushOutput();
        }

        private void FlushOutput() {
            if (_outputCount == 0) return;
            byte[] output = _output ?? throw new ObjectDisposedException(nameof(StreamLsbBitWriter));
            _destination.Write(output, 0, _outputCount);
            _outputCount = 0;
        }
    }
}
