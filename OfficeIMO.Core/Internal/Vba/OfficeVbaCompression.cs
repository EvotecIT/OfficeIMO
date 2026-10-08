using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeIMO.Core.Internal;

/// <summary>Bounded MS-OVBA source compression shared by editing and signature binding.</summary>
internal static class OfficeVbaCompression {
    /// <summary>Writes native-compatible MS-OVBA literal/copy chunks without changing source bytes.</summary>
    internal static byte[] Compress(byte[] source) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        using var output = new MemoryStream();
        output.WriteByte(1);
        int position = 0;
        while (position < source.Length) {
            int count = Math.Min(4096, source.Length - position);
            byte[] compressed = CompressChunk(source, position, count);
            if (compressed.Length > 4096) {
                // Keep newly written source token-compressed for native Office compatibility.
                // 3640 literals plus 455 flag bytes fit without padding the source.
                count = Math.Min(count, 3640);
                compressed = CompressChunk(source, position, count);
            }
            ushort header = (ushort)(0xb000 | (compressed.Length - 1));
            output.WriteByte((byte)header);
            output.WriteByte((byte)(header >> 8));
            output.Write(compressed, 0, compressed.Length);
            position += count;
        }
        return output.ToArray();
    }

    /// <summary>Checks chunk flags in a module container already validated by project loading.</summary>
    internal static bool ContainsRawChunk(byte[] validatedModuleStream, int sourceOffset) {
        for (int position = sourceOffset + 1; position < validatedModuleStream.Length;) {
            ushort header = ReadUInt16(validatedModuleStream, position);
            if ((header & 0x8000) == 0) return true;
            position += (header & 0x0fff) + 3;
        }
        return false;
    }

    private static byte[] CompressChunk(byte[] source, int start, int count) {
        // Each dictionary and search is confined to one 4 KiB chunk. A bounded hash chain
        // avoids quadratic work on large inputs while allowing overlapping copy tokens.
        var heads = new Dictionary<int, int>();
        var previous = new int[count];
        using var output = new MemoryStream();
        int position = 0;
        while (position < count) {
            long flagPosition = output.Position;
            output.WriteByte(0);
            byte flags = 0;
            for (int bit = 0; bit < 8 && position < count; bit++) {
                int bitCount = 4;
                while (bitCount < 12 && (1 << bitCount) < position) bitCount++;
                int maximumLength = Math.Min(count - position, (0xffff >> bitCount) + 3);
                int bestLength = 0, bestOffset = 0;
                if (position + 2 < count && heads.TryGetValue(Key(source, start + position), out int candidate)) {
                    for (int attempt = 0; candidate >= 0 && attempt < 64; attempt++, candidate = previous[candidate]) {
                        int length = 0;
                        while (length < maximumLength && source[start + candidate + length] == source[start + position + length]) length++;
                        if (length > bestLength) { bestLength = length; bestOffset = position - candidate; }
                        if (bestLength == maximumLength) break;
                    }
                }
                int consumed = bestLength >= 3 ? bestLength : 1;
                if (bestLength >= 3) {
                    flags |= (byte)(1 << bit);
                    ushort token = (ushort)(((bestOffset - 1) << (16 - bitCount)) | (bestLength - 3));
                    output.WriteByte((byte)token);
                    output.WriteByte((byte)(token >> 8));
                } else output.WriteByte(source[start + position]);
                for (int inserted = position; inserted < position + consumed && inserted + 2 < count; inserted++) {
                    int key = Key(source, start + inserted);
                    previous[inserted] = heads.TryGetValue(key, out int prior) ? prior : -1;
                    heads[key] = inserted;
                }
                position += consumed;
            }
            long end = output.Position;
            output.Position = flagPosition;
            output.WriteByte(flags);
            output.Position = end;
        }
        return output.ToArray();
    }

    private static int Key(byte[] source, int offset) => source[offset] | source[offset + 1] << 8 | source[offset + 2] << 16;

    internal static bool TryDecompress(byte[] input, int maximumOutputBytes,
        out byte[] output, out string detail) {
        output = Array.Empty<byte>();
        if (input.Length == 0 || input[0] != 0x01) {
            detail = "The MS-OVBA compressed container signature is missing.";
            return false;
        }
        var decompressed = new List<byte>(Math.Min(input.Length * 2, maximumOutputBytes));
        int position = 1;
        while (position < input.Length) {
            int headerPosition = position;
            if (!TryReadUInt16(input, ref position, out ushort header)) {
                detail = "The compressed container ends inside a chunk header.";
                return false;
            }
            int chunkSize = (header & 0x0FFF) + 3;
            int chunkEnd = headerPosition + chunkSize;
            if ((header & 0x7000) != 0x3000 || chunkEnd < position || chunkEnd > input.Length) {
                detail = "The compressed container has an invalid chunk header.";
                return false;
            }
            int chunkOutputStart = decompressed.Count;
            if ((header & 0x8000) == 0) {
                if (chunkSize != 4098 || chunkEnd - position != 4096 ||
                    decompressed.Count > maximumOutputBytes - 4096) {
                    detail = "The compressed container has an invalid or oversized raw chunk.";
                    return false;
                }
                for (; position < chunkEnd; position++) decompressed.Add(input[position]);
                continue;
            }
            while (position < chunkEnd) {
                byte flags = input[position++];
                for (int bit = 0; bit < 8 && position < chunkEnd; bit++) {
                    if ((flags & 1 << bit) == 0) {
                        if (decompressed.Count >= maximumOutputBytes) {
                            detail = "The expanded MS-OVBA container exceeds the configured byte limit.";
                            return false;
                        }
                        decompressed.Add(input[position++]);
                        continue;
                    }
                    if (!TryReadUInt16(input, ref position, out ushort token) || position > chunkEnd) {
                        detail = "The compressed container ends inside a copy token.";
                        return false;
                    }
                    int decompressedPosition = decompressed.Count - chunkOutputStart;
                    int bitCount = 4;
                    while (bitCount < 12 && 1 << bitCount < decompressedPosition) bitCount++;
                    int lengthMask = 0xFFFF >> bitCount;
                    int offset = ((token & ~lengthMask) >> (16 - bitCount)) + 1;
                    int length = (token & lengthMask) + 3;
                    int sourceOffset = decompressed.Count - offset;
                    if (decompressedPosition <= 0 || sourceOffset < chunkOutputStart ||
                        decompressedPosition + length > 4096 || decompressed.Count > maximumOutputBytes - length) {
                        detail = "The compressed container has an out-of-range copy token.";
                        return false;
                    }
                    for (int copied = 0; copied < length; copied++) decompressed.Add(decompressed[sourceOffset + copied]);
                }
            }
        }
        output = decompressed.ToArray();
        detail = string.Empty;
        return true;
    }

    private static ushort ReadUInt16(byte[] bytes, int offset) => (ushort)(bytes[offset] | bytes[offset + 1] << 8);
    private static bool TryReadUInt16(byte[] bytes, ref int position, out ushort value) { value = 0; if (position < 0 || position > bytes.Length - 2) return false; value = ReadUInt16(bytes, position); position += 2; return true; }
}
