using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Validates the required headers and compressed directory in a VBA project storage.</summary>
    internal static class LegacyPptVbaProjectValidator {
        private const int MaximumDirectoryLength = 16 * 1024 * 1024;

        internal static bool TryValidate(
            IReadOnlyDictionary<string, byte[]> streams, out string? reason) {
            if (!streams.TryGetValue("VBA/_VBA_PROJECT", out byte[]? project)
                || project.Length < 7
                || project[0] != 0xCC
                || project[1] != 0x61
                || project[4] != 0x00) {
                reason = "The VBA/_VBA_PROJECT stream has an invalid MS-OVBA header.";
                return false;
            }
            if (!streams.TryGetValue("VBA/dir", out byte[]? directory)) {
                reason = "The compound storage has no VBA/dir stream.";
                return false;
            }
            if (!TryDecompressDirectory(directory, out byte[] decompressed,
                    out reason)) {
                return false;
            }
            return TryParseDirectory(decompressed, out reason);
        }

        private static bool TryDecompressDirectory(byte[] input,
            out byte[] output, out string? reason) {
            var decompressed = new List<byte>(Math.Min(input.Length * 2,
                MaximumDirectoryLength));
            output = Array.Empty<byte>();
            if (input.Length == 0 || input[0] != 0x01) {
                reason = "The VBA/dir stream has no MS-OVBA compressed-container signature.";
                return false;
            }

            int position = 1;
            while (position < input.Length) {
                int headerPosition = position;
                if (!TryReadUInt16(input, ref position, out ushort header)) {
                    reason = "The VBA/dir stream ends inside a compressed-chunk header.";
                    return false;
                }
                int chunkSize = (header & 0x0FFF) + 3;
                int chunkEnd = headerPosition + chunkSize;
                if ((header & 0x7000) != 0x3000
                    || chunkEnd < position || chunkEnd > input.Length) {
                    reason = "The VBA/dir stream contains an invalid compressed-chunk header.";
                    return false;
                }

                int chunkOutputStart = decompressed.Count;
                bool compressed = (header & 0x8000) != 0;
                if (!compressed) {
                    if (chunkSize != 4098 || chunkEnd - position != 4096
                        || !TryAppendBytes(input, position, 4096, decompressed)) {
                        reason = "The VBA/dir stream contains an invalid raw chunk.";
                        return false;
                    }
                    position = chunkEnd;
                    continue;
                }

                while (position < chunkEnd) {
                    byte flags = input[position++];
                    for (int bit = 0; bit < 8 && position < chunkEnd; bit++) {
                        if ((flags & (1 << bit)) == 0) {
                            if (!TryAppendByte(input[position++], decompressed)) {
                                reason = "The expanded VBA/dir stream exceeds the supported limit.";
                                return false;
                            }
                            continue;
                        }

                        if (!TryReadUInt16(input, ref position, out ushort token)
                            || position > chunkEnd) {
                            reason = "The VBA/dir stream ends inside a copy token.";
                            return false;
                        }
                        int decompressedPosition = decompressed.Count - chunkOutputStart;
                        if (decompressedPosition <= 0) {
                            reason = "The VBA/dir stream starts a chunk with an invalid copy token.";
                            return false;
                        }
                        int bitCount = 4;
                        while (bitCount < 12 && (1 << bitCount) < decompressedPosition) {
                            bitCount++;
                        }
                        int lengthMask = 0xFFFF >> bitCount;
                        int offset = ((token & ~lengthMask) >> (16 - bitCount)) + 1;
                        int length = (token & lengthMask) + 3;
                        int source = decompressed.Count - offset;
                        if (source < chunkOutputStart
                            || decompressedPosition + length > 4096) {
                            reason = "The VBA/dir stream contains an out-of-range copy token.";
                            return false;
                        }
                        for (int copied = 0; copied < length; copied++) {
                            if (!TryAppendByte(decompressed[source + copied], decompressed)) {
                                reason = "The expanded VBA/dir stream exceeds the supported limit.";
                                return false;
                            }
                        }
                    }
                }
                if (position != chunkEnd) {
                    reason = "The VBA/dir compressed chunk has an invalid boundary.";
                    return false;
                }
            }

            output = decompressed.ToArray();
            reason = null;
            return true;
        }

        private static bool TryParseDirectory(byte[] directory,
            out string? reason) {
            int position = 0;
            if (!TryReadSizedRecord(directory, ref position, 0x0001, 4, false)
                || !TryReadSizedRecord(directory, ref position, 0x0002, 4, false)) {
                reason = "The expanded VBA/dir stream has invalid project-system or locale records.";
                return false;
            }
            if (PeekUInt16(directory, position) == 0x0014
                && !TryReadSizedRecord(directory, ref position, 0x0014, 4, false)) {
                reason = "The expanded VBA/dir stream has an invalid invocation-locale record.";
                return false;
            }
            if (!TryReadSizedRecord(directory, ref position, 0x0003, 2, false)
                || !TryReadSizedRecord(directory, ref position, 0x0004, null, true)) {
                reason = "The expanded VBA/dir stream has invalid code-page or project-name records.";
                return false;
            }
            int terminator = directory.Length - 6;
            if (terminator < position || PeekUInt16(directory, terminator) != 0x0010
                || directory[terminator + 2] != 0
                || directory[terminator + 3] != 0
                || directory[terminator + 4] != 0
                || directory[terminator + 5] != 0) {
                reason = "The expanded VBA/dir stream has no valid project terminator.";
                return false;
            }
            reason = null;
            return true;
        }

        private static bool TryReadSizedRecord(byte[] bytes, ref int position,
            ushort expectedId, int? expectedSize, bool requireData) {
            if (!TryReadUInt16(bytes, ref position, out ushort id)
                || id != expectedId || !TryReadUInt32(bytes, ref position, out uint size)
                || size > int.MaxValue || expectedSize.HasValue && size != expectedSize.Value
                || requireData && size == 0 || position + (long)size > bytes.Length) {
                return false;
            }
            position += (int)size;
            return true;
        }

        private static ushort PeekUInt16(byte[] bytes, int position) =>
            position >= 0 && position + 2 <= bytes.Length
                ? (ushort)(bytes[position] | bytes[position + 1] << 8)
                : ushort.MaxValue;

        private static bool TryReadUInt16(byte[] bytes, ref int position,
            out ushort value) {
            if (position < 0 || position + 2 > bytes.Length) {
                value = 0;
                return false;
            }
            value = (ushort)(bytes[position] | bytes[position + 1] << 8);
            position += 2;
            return true;
        }

        private static bool TryReadUInt32(byte[] bytes, ref int position,
            out uint value) {
            if (position < 0 || position + 4 > bytes.Length) {
                value = 0;
                return false;
            }
            value = (uint)(bytes[position]
                | bytes[position + 1] << 8
                | bytes[position + 2] << 16
                | bytes[position + 3] << 24);
            position += 4;
            return true;
        }

        private static bool TryAppendBytes(byte[] input, int offset, int count,
            List<byte> output) {
            if (count < 0 || offset < 0 || offset + count > input.Length
                || output.Count > MaximumDirectoryLength - count) {
                return false;
            }
            for (int index = 0; index < count; index++) {
                output.Add(input[offset + index]);
            }
            return true;
        }

        private static bool TryAppendByte(byte value, List<byte> output) {
            if (output.Count >= MaximumDirectoryLength) {
                return false;
            }
            output.Add(value);
            return true;
        }
    }
}
