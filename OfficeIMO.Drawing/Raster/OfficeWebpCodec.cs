using System;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free lossless WebP encoder for RGBA images.
/// </summary>
/// <remarks>
/// The encoder intentionally uses a deterministic literal-only VP8L stream. It favors a small,
/// auditable implementation over compression efficiency while producing standards-compatible WebP.
/// </remarks>
public static class OfficeWebpCodec {
    private const double MaximumDpi = 1000000D;

    /// <summary>Returns whether the payload is a RIFF WebP container.</summary>
    public static bool IsWebp(byte[]? encodedBytes) =>
        encodedBytes != null && encodedBytes.Length >= 12 &&
        HasAscii(encodedBytes, 0, "RIFF") && HasAscii(encodedBytes, 8, "WEBP");

    /// <summary>Encodes an RGBA image as a lossless VP8L WebP image.</summary>
    public static byte[] Encode(OfficeRasterImage image) {
        return EncodeCore(image, includeResolutionMetadata: false, 96D, 96D);
    }

    /// <summary>Encodes an RGBA image as a lossless VP8L WebP image with Exif resolution metadata.</summary>
    public static byte[] Encode(OfficeRasterImage image, double dpiX, double dpiY) {
        ValidateDpi(dpiX, nameof(dpiX));
        ValidateDpi(dpiY, nameof(dpiY));
        return EncodeCore(image, includeResolutionMetadata: true, dpiX, dpiY);
    }

    private static byte[] EncodeCore(
        OfficeRasterImage image,
        bool includeResolutionMetadata,
        double dpiX,
        double dpiY) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        if (image.Width > OfficeRasterImageEncoder.WebpMaximumDimension) throw new ArgumentOutOfRangeException(nameof(image), "WebP width cannot exceed 16,384 pixels.");
        if (image.Height > OfficeRasterImageEncoder.WebpMaximumDimension) throw new ArgumentOutOfRangeException(nameof(image), "WebP height cannot exceed 16,384 pixels.");

        byte[] pixels = image.GetPixels();
        bool hasAlpha = HasTransparency(pixels);
        byte[] payload;
        using (var stream = new MemoryStream()) {
            stream.WriteByte(0x2F);
            var writer = new LsbBitWriter(stream);
            writer.WriteBits((uint)(image.Width - 1), 14);
            writer.WriteBits((uint)(image.Height - 1), 14);
            writer.WriteBits(hasAlpha ? 1U : 0U, 1);
            writer.WriteBits(0, 3);

            writer.WriteBits(0, 1); // no transforms
            writer.WriteBits(0, 1); // no color cache
            writer.WriteBits(0, 1); // one Huffman code group for the whole image

            WriteLiteralTree(writer, 280);
            WriteLiteralTree(writer, 256);
            WriteLiteralTree(writer, 256);
            WriteLiteralTree(writer, 256);
            WriteSingleSymbolTree(writer);

            for (int offset = 0; offset < pixels.Length; offset += 4) {
                writer.WriteBits(ReverseByte(pixels[offset + 1]), 8); // green
                writer.WriteBits(ReverseByte(pixels[offset]), 8);     // red
                writer.WriteBits(ReverseByte(pixels[offset + 2]), 8); // blue
                writer.WriteBits(ReverseByte(pixels[offset + 3]), 8); // alpha
            }

            writer.Flush();
            payload = stream.ToArray();
        }

        return includeResolutionMetadata
            ? BuildExtendedContainer(image.Width, image.Height, hasAlpha, payload, dpiX, dpiY)
            : BuildSimpleContainer(payload);
    }

    private static byte[] BuildSimpleContainer(byte[] payload) {
        int paddedPayloadLength = checked(payload.Length + (payload.Length & 1));
        int fileLength = checked(20 + paddedPayloadLength);
        byte[] output = new byte[fileLength];
        WriteAscii(output, 0, "RIFF");
        WriteUInt32(output, 4, fileLength - 8);
        WriteAscii(output, 8, "WEBP");
        WriteAscii(output, 12, "VP8L");
        WriteUInt32(output, 16, payload.Length);
        Buffer.BlockCopy(payload, 0, output, 20, payload.Length);
        return output;
    }

    private static byte[] BuildExtendedContainer(
        int width,
        int height,
        bool hasAlpha,
        byte[] payload,
        double dpiX,
        double dpiY) {
        byte[] exif = CreateResolutionExif(width, height, dpiX, dpiY);
        int paddedPayloadLength = checked(payload.Length + (payload.Length & 1));
        int paddedExifLength = checked(exif.Length + (exif.Length & 1));
        int fileLength = checked(12 + 18 + 8 + paddedPayloadLength + 8 + paddedExifLength);
        byte[] output = new byte[fileLength];
        WriteAscii(output, 0, "RIFF");
        WriteUInt32(output, 4, fileLength - 8);
        WriteAscii(output, 8, "WEBP");

        WriteAscii(output, 12, "VP8X");
        WriteUInt32(output, 16, 10);
        output[20] = (byte)(0x08 | (hasAlpha ? 0x10 : 0x00));
        WriteUInt24(output, 24, width - 1);
        WriteUInt24(output, 27, height - 1);

        const int vp8lChunkOffset = 30;
        WriteAscii(output, vp8lChunkOffset, "VP8L");
        WriteUInt32(output, vp8lChunkOffset + 4, payload.Length);
        Buffer.BlockCopy(payload, 0, output, vp8lChunkOffset + 8, payload.Length);

        int exifChunkOffset = vp8lChunkOffset + 8 + paddedPayloadLength;
        WriteAscii(output, exifChunkOffset, "EXIF");
        WriteUInt32(output, exifChunkOffset + 4, exif.Length);
        Buffer.BlockCopy(exif, 0, output, exifChunkOffset + 8, exif.Length);
        return output;
    }

    private static byte[] CreateResolutionExif(int width, int height, double dpiX, double dpiY) {
        const int entryCount = 5;
        const int ifdOffset = 8;
        const int xResolutionOffset = ifdOffset + 2 + entryCount * 12 + 4;
        const int yResolutionOffset = xResolutionOffset + 8;
        byte[] output = new byte[yResolutionOffset + 8];

        output[0] = (byte)'I';
        output[1] = (byte)'I';
        WriteUInt16(output, 2, 42);
        WriteUInt32(output, 4, ifdOffset);
        WriteUInt16(output, ifdOffset, entryCount);
        WriteExifEntry(output, ifdOffset + 2, 256, 4, 1, width);
        WriteExifEntry(output, ifdOffset + 14, 257, 4, 1, height);
        WriteExifEntry(output, ifdOffset + 26, 282, 5, 1, xResolutionOffset);
        WriteExifEntry(output, ifdOffset + 38, 283, 5, 1, yResolutionOffset);
        WriteExifEntry(output, ifdOffset + 50, 296, 3, 1, 2);
        WriteRational(output, xResolutionOffset, dpiX);
        WriteRational(output, yResolutionOffset, dpiY);
        return output;
    }

    private static void WriteExifEntry(
        byte[] output,
        int offset,
        int tag,
        int type,
        int count,
        int valueOrOffset) {
        WriteUInt16(output, offset, tag);
        WriteUInt16(output, offset + 2, type);
        WriteUInt32(output, offset + 4, count);
        WriteUInt32(output, offset + 8, valueOrOffset);
    }

    private static void WriteRational(byte[] output, int offset, double value) {
        uint denominator = (uint)Math.Max(1D, Math.Min(10000D, Math.Floor(uint.MaxValue / value)));
        uint numerator = checked((uint)Math.Round(value * denominator, MidpointRounding.AwayFromZero));
        WriteUInt32(output, offset, numerator);
        WriteUInt32(output, offset + 4, denominator);
    }

    /// <summary>
    /// Attempts to decode the deterministic literal-only VP8L subset emitted by
    /// <see cref="Encode(OfficeRasterImage)"/> and its resolution-aware overload.
    /// General VP8/VP8L features remain the responsibility of an optional caller codec.
    /// </summary>
    public static bool TryDecode(byte[]? encodedBytes, out OfficeRasterImage? image) {
        image = null;
        if (!IsWebp(encodedBytes) || encodedBytes == null ||
            encodedBytes.Length < 22 ||
            encodedBytes.Length > OfficeRasterGuards.MaximumEncodedBytes) {
            return false;
        }

        try {
            int riffLength = ReadUInt32(encodedBytes, 4);
            if (riffLength != encodedBytes.Length - 8 ||
                !TryFindChunk(encodedBytes, "VP8L", out int payloadOffset, out int payloadLength) ||
                payloadLength < 5 ||
                encodedBytes[payloadOffset] != 0x2F) {
                return false;
            }

            var reader = new LsbBitReader(encodedBytes, payloadOffset + 1, payloadLength - 1);
            int width = checked((int)reader.ReadBits(14) + 1);
            int height = checked((int)reader.ReadBits(14) + 1);
            reader.ReadBits(1); // alpha hint
            if (reader.ReadBits(3) != 0 ||
                reader.ReadBits(1) != 0 || // no transforms
                reader.ReadBits(1) != 0 || // no color cache
                reader.ReadBits(1) != 0) { // one Huffman group
                return false;
            }
            if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out int pixels) ||
                !TryReadLiteralTree(reader, 280) ||
                !TryReadLiteralTree(reader, 256) ||
                !TryReadLiteralTree(reader, 256) ||
                !TryReadLiteralTree(reader, 256) ||
                !TryReadSingleSymbolTree(reader)) {
                return false;
            }

            if (!reader.HasBits((long)pixels * 32L)) {
                return false;
            }

            byte[] rgba = OfficeRasterGuards.AllocateRgba32(width, height, "WebP decoded pixels exceed the managed limit.");
            for (int pixel = 0; pixel < pixels; pixel++) {
                int offset = pixel * 4;
                rgba[offset + 1] = (byte)ReverseByte((byte)reader.ReadBits(8));
                rgba[offset] = (byte)ReverseByte((byte)reader.ReadBits(8));
                rgba[offset + 2] = (byte)ReverseByte((byte)reader.ReadBits(8));
                rgba[offset + 3] = (byte)ReverseByte((byte)reader.ReadBits(8));
            }
            if (!reader.HasOnlyZeroPadding()) {
                return false;
            }
            image = OfficeRasterImage.FromRgba32(width, height, rgba);
            return true;
        } catch (FormatException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool TryFindChunk(
        byte[] input,
        string chunkType,
        out int payloadOffset,
        out int payloadLength) {
        payloadOffset = 0;
        payloadLength = 0;
        int offset = 12;
        while (offset < input.Length) {
            if (offset > input.Length - 8) return false;
            int chunkLength = ReadUInt32(input, offset + 4);
            int chunkPayloadOffset = checked(offset + 8);
            int chunkEnd = checked(chunkPayloadOffset + chunkLength);
            int paddedChunkEnd = checked(chunkEnd + (chunkLength & 1));
            if (chunkEnd > input.Length ||
                paddedChunkEnd > input.Length ||
                (chunkLength & 1) != 0 && input[chunkEnd] != 0) {
                return false;
            }

            if (HasAscii(input, offset, chunkType)) {
                if (payloadOffset != 0) return false;
                payloadOffset = chunkPayloadOffset;
                payloadLength = chunkLength;
            }

            offset = paddedChunkEnd;
        }

        return offset == input.Length && payloadOffset != 0;
    }

    private static void WriteLiteralTree(LsbBitWriter writer, int alphabetSize) {
        writer.WriteBits(0, 1); // normal Huffman tree
        writer.WriteBits(8, 4); // store 12 code-length-code entries

        int[] codeLengthDepths = { 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1 };
        for (int i = 0; i < codeLengthDepths.Length; i++) {
            writer.WriteBits((uint)codeLengthDepths[i], 3);
        }

        writer.WriteBits(0, 1); // use the full alphabet length
        for (int i = 0; i < 256; i++) {
            writer.WriteBits(0, 1); // code-length symbol 8
        }

        if (alphabetSize == 280) {
            writer.WriteBits(1, 1);  // code-length symbol 18
            writer.WriteBits(13, 7); // 24 zero lengths (11 + 13)
        }
    }

    private static void WriteSingleSymbolTree(LsbBitWriter writer) {
        writer.WriteBits(1, 1); // small tree
        writer.WriteBits(0, 1); // one symbol
        writer.WriteBits(0, 1); // symbol uses one bit
        writer.WriteBits(0, 1); // symbol zero
    }

    private static bool TryReadLiteralTree(LsbBitReader reader, int alphabetSize) {
        if (reader.ReadBits(1) != 0 || reader.ReadBits(4) != 8) return false;
        int[] codeLengthDepths = { 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1 };
        for (int index = 0; index < codeLengthDepths.Length; index++) {
            if (reader.ReadBits(3) != (uint)codeLengthDepths[index]) return false;
        }
        if (reader.ReadBits(1) != 0) return false;
        for (int index = 0; index < 256; index++) {
            if (reader.ReadBits(1) != 0) return false;
        }
        return alphabetSize != 280 ||
               reader.ReadBits(1) == 1 && reader.ReadBits(7) == 13;
    }

    private static bool TryReadSingleSymbolTree(LsbBitReader reader) =>
        reader.ReadBits(1) == 1 &&
        reader.ReadBits(1) == 0 &&
        reader.ReadBits(1) == 0 &&
        reader.ReadBits(1) == 0;

    private static uint ReverseByte(byte value) {
        uint reversed = value;
        reversed = ((reversed & 0x55U) << 1) | ((reversed >> 1) & 0x55U);
        reversed = ((reversed & 0x33U) << 2) | ((reversed >> 2) & 0x33U);
        return ((reversed & 0x0FU) << 4) | ((reversed >> 4) & 0x0FU);
    }

    private static bool HasTransparency(byte[] pixels) {
        for (int i = 3; i < pixels.Length; i += 4) {
            if (pixels[i] != 255) return true;
        }

        return false;
    }

    private static bool HasAscii(byte[] data, int offset, string value) {
        if (offset < 0 || offset + value.Length > data.Length) return false;
        for (int i = 0; i < value.Length; i++) {
            if (data[offset + i] != (byte)value[i]) return false;
        }

        return true;
    }

    private static void WriteAscii(byte[] output, int offset, string value) {
        for (int i = 0; i < value.Length; i++) output[offset + i] = (byte)value[i];
    }

    private static void WriteUInt32(byte[] output, int offset, int value) {
        output[offset] = (byte)value;
        output[offset + 1] = (byte)(value >> 8);
        output[offset + 2] = (byte)(value >> 16);
        output[offset + 3] = (byte)(value >> 24);
    }

    private static void WriteUInt32(byte[] output, int offset, uint value) {
        output[offset] = (byte)value;
        output[offset + 1] = (byte)(value >> 8);
        output[offset + 2] = (byte)(value >> 16);
        output[offset + 3] = (byte)(value >> 24);
    }

    private static void WriteUInt16(byte[] output, int offset, int value) {
        output[offset] = (byte)value;
        output[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt24(byte[] output, int offset, int value) {
        output[offset] = (byte)value;
        output[offset + 1] = (byte)(value >> 8);
        output[offset + 2] = (byte)(value >> 16);
    }

    private static void ValidateDpi(double dpi, string name) {
        if (dpi <= 0D || double.IsNaN(dpi) || double.IsInfinity(dpi) || dpi > MaximumDpi) {
            throw new ArgumentOutOfRangeException(name, "WebP DPI must be finite, positive, and no greater than 1,000,000.");
        }
    }

    private static int ReadUInt32(byte[] input, int offset) {
        if (offset < 0 || offset > input.Length - 4) throw new FormatException("WebP integer field is truncated.");
        uint value = (uint)(input[offset] |
                            input[offset + 1] << 8 |
                            input[offset + 2] << 16 |
                            input[offset + 3] << 24);
        if (value > int.MaxValue) throw new FormatException("WebP length exceeds supported integer bounds.");
        return (int)value;
    }

    private sealed class LsbBitWriter {
        private readonly Stream _stream;
        private ulong _buffer;
        private int _bitCount;

        public LsbBitWriter(Stream stream) {
            _stream = stream;
        }

        public void WriteBits(uint value, int count) {
            if (count < 0 || count > 32) throw new ArgumentOutOfRangeException(nameof(count));
            ulong mask = count == 32 ? uint.MaxValue : ((1UL << count) - 1UL);
            _buffer |= ((ulong)value & mask) << _bitCount;
            _bitCount += count;
            while (_bitCount >= 8) {
                _stream.WriteByte((byte)_buffer);
                _buffer >>= 8;
                _bitCount -= 8;
            }
        }

        public void Flush() {
            if (_bitCount > 0) {
                _stream.WriteByte((byte)_buffer);
                _buffer = 0;
                _bitCount = 0;
            }
        }
    }

    private sealed class LsbBitReader {
        private readonly byte[] _input;
        private readonly int _end;
        private int _offset;
        private ulong _buffer;
        private int _bitCount;

        internal LsbBitReader(byte[] input, int offset, int count) {
            _input = input;
            _offset = offset;
            _end = checked(offset + count);
            if (offset < 0 || count < 0 || _end > input.Length) {
                throw new FormatException("WebP bitstream is truncated.");
            }
        }

        internal uint ReadBits(int count) {
            if (count < 0 || count > 32) throw new FormatException("WebP bit count is invalid.");
            while (_bitCount < count) {
                if (_offset >= _end) throw new FormatException("WebP bitstream is truncated.");
                _buffer |= (ulong)_input[_offset++] << _bitCount;
                _bitCount += 8;
            }
            ulong mask = count == 32 ? uint.MaxValue : (1UL << count) - 1UL;
            uint value = (uint)(_buffer & mask);
            _buffer >>= count;
            _bitCount -= count;
            return value;
        }

        internal bool HasBits(long count) =>
            count >= 0L &&
            count <= _bitCount + ((long)_end - _offset) * 8L;

        internal bool HasOnlyZeroPadding() {
            if (_bitCount > 0) {
                ulong mask = (1UL << _bitCount) - 1UL;
                if ((_buffer & mask) != 0UL) return false;
            }
            while (_offset < _end) {
                if (_input[_offset++] != 0) return false;
            }
            return true;
        }
    }
}
