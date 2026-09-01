namespace OfficeIMO.IWork.Internal;

internal static class IWorkSnappy {
    internal static byte[] DecodeIwa(byte[] data, IWorkReadOptions options, long remainingTotalBytes) {
        if (data.Length > options.MaximumIwaBytes) {
            throw new InvalidDataException($"IWA entry size {data.Length} exceeds the configured limit of {options.MaximumIwaBytes} bytes.");
        }
        if (remainingTotalBytes <= 0) {
            throw new InvalidDataException($"Combined decompressed IWA data exceeds the configured limit of {options.MaximumTotalDecompressedIwaBytes} bytes.");
        }

        int maximumOutputBytes = (int)Math.Min(options.MaximumDecompressedIwaBytes,
            Math.Min(remainingTotalBytes, int.MaxValue));

        using var output = new MemoryStream(Math.Min(data.Length, maximumOutputBytes));
        int offset = 0;
        while (offset < data.Length) {
            if (data.Length - offset < 4) throw new InvalidDataException($"Truncated IWA chunk header at offset {offset}.");
            byte chunkType = data[offset];
            int length = data[offset + 1] | data[offset + 2] << 8 | data[offset + 3] << 16;
            offset += 4;
            if (length < 0 || offset > data.Length - length) throw new InvalidDataException($"Truncated IWA chunk payload at offset {offset}.");
            if (chunkType != 0) throw new InvalidDataException($"Unsupported IWA chunk type 0x{chunkType:X2} at offset {offset - 4}.");

            int remainingEntryBytes = maximumOutputBytes - checked((int)output.Length);
            byte[] block = DecodeRaw(data, offset, length,
                Math.Min(options.MaximumSnappyChunkBytes, remainingEntryBytes));
            if (output.Length > maximumOutputBytes - block.Length) {
                throw new InvalidDataException($"Decompressed IWA data exceeds the applicable entry or source-wide limit of {maximumOutputBytes} bytes.");
            }
            output.Write(block, 0, block.Length);
            offset += length;
        }
        return output.ToArray();
    }

    private static byte[] DecodeRaw(byte[] input, int start, int length, int maximumOutputBytes) {
        int offset = start;
        int end = start + length;
        ulong declared = ReadVarint(input, ref offset, end);
        if (declared > (ulong)maximumOutputBytes || declared > int.MaxValue) {
            throw new InvalidDataException($"Snappy block declares {declared} bytes, above the applicable entry or source-wide limit of {maximumOutputBytes}.");
        }

        var output = new byte[(int)declared];
        int written = 0;
        while (offset < end && written < output.Length) {
            byte tag = input[offset++];
            int kind = tag & 3;
            if (kind == 0) {
                int literalCode = tag >> 2;
                int literalLength;
                if (literalCode < 60) {
                    literalLength = literalCode + 1;
                } else {
                    int byteCount = literalCode - 59;
                    if (byteCount < 1 || byteCount > 4 || offset > end - byteCount) {
                        throw new InvalidDataException("Invalid Snappy literal length.");
                    }
                    uint encoded = 0;
                    for (int index = 0; index < byteCount; index++) encoded |= (uint)input[offset++] << (index * 8);
                    if (encoded == uint.MaxValue) throw new InvalidDataException("Snappy literal length overflows the supported range.");
                    literalLength = checked((int)encoded + 1);
                }
                EnsureCopyBounds(offset, literalLength, end, written, output.Length, "literal");
                Buffer.BlockCopy(input, offset, output, written, literalLength);
                offset += literalLength;
                written += literalLength;
                continue;
            }

            int copyLength;
            int copyOffset;
            if (kind == 1) {
                if (offset >= end) throw new InvalidDataException("Truncated Snappy copy-1 tag.");
                copyLength = 4 + ((tag >> 2) & 7);
                copyOffset = ((tag & 0xe0) << 3) | input[offset++];
            } else if (kind == 2) {
                if (offset > end - 2) throw new InvalidDataException("Truncated Snappy copy-2 tag.");
                copyLength = 1 + (tag >> 2);
                copyOffset = input[offset] | input[offset + 1] << 8;
                offset += 2;
            } else {
                if (offset > end - 4) throw new InvalidDataException("Truncated Snappy copy-4 tag.");
                copyLength = 1 + (tag >> 2);
                uint rawOffset = IWorkProtobuf.ReadUInt32(input, offset);
                if (rawOffset > int.MaxValue) throw new InvalidDataException("Snappy copy offset exceeds the supported range.");
                copyOffset = (int)rawOffset;
                offset += 4;
            }

            if (copyOffset <= 0 || copyOffset > written || copyLength > output.Length - written) {
                throw new InvalidDataException("Snappy copy references data outside the decoded prefix.");
            }
            for (int index = 0; index < copyLength; index++) output[written + index] = output[written - copyOffset + index];
            written += copyLength;
        }

        if (offset != end || written != output.Length) {
            throw new InvalidDataException($"Snappy block decoded {written} of {output.Length} declared bytes.");
        }
        return output;
    }

    private static ulong ReadVarint(byte[] input, ref int offset, int end) {
        ulong result = 0;
        int shift = 0;
        while (shift < 64) {
            if (offset >= end) throw new InvalidDataException("Truncated Snappy length varint.");
            byte value = input[offset++];
            if (shift == 63 && (value & 0xfe) != 0) throw new InvalidDataException("Snappy length varint overflows UInt64.");
            result |= (ulong)(value & 0x7f) << shift;
            if ((value & 0x80) == 0) return result;
            shift += 7;
        }
        throw new InvalidDataException("Snappy length varint exceeds ten bytes.");
    }

    private static void EnsureCopyBounds(int sourceOffset, int length, int sourceEnd,
        int targetOffset, int targetLength, string operation) {
        if (length < 0 || sourceOffset > sourceEnd - length || targetOffset > targetLength - length) {
            throw new InvalidDataException($"Snappy {operation} exceeds the source or declared output bounds.");
        }
    }
}