using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Validates the required headers and compressed directory in a VBA project storage.</summary>
    internal static class LegacyPptVbaProjectValidator {
        private const int MaximumDirectoryLength = 16 * 1024 * 1024;
        private const int MaximumModuleCount = 1024;
        private const long MaximumModuleValidationBytes = 64L * 1024L * 1024L;

        static LegacyPptVbaProjectValidator() {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        internal static bool TryValidate(
            IReadOnlyDictionary<string, byte[]> streams, out string? reason) {
            if (!streams.TryGetValue("PROJECT", out byte[]? projectMetadata)
                || projectMetadata.Length == 0
                || projectMetadata.Length > MaximumDirectoryLength
                || !TryValidateProjectMetadata(projectMetadata)) {
                reason = "The compound storage has no valid mandatory PROJECT stream.";
                return false;
            }
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
            return TryParseDirectory(decompressed, streams, out reason);
        }

        private static bool TryValidateProjectMetadata(byte[] project) {
            string text = Encoding.ASCII.GetString(project);
            string[] lines = text.Split(new[] { '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries);
            string? id = TryGetProjectRecordValue(lines, "ID");
            string? name = TryGetProjectRecordValue(lines, "Name");
            return Guid.TryParse(id, out _)
                && !string.IsNullOrWhiteSpace(name);
        }

        private static string? TryGetProjectRecordValue(
            IEnumerable<string> lines, string key) {
            string prefix = key + "=";
            string? raw = lines.FirstOrDefault(line => line.StartsWith(
                prefix, StringComparison.OrdinalIgnoreCase));
            if (raw == null) return null;
            string value = raw.Substring(prefix.Length).Trim();
            if (value.Length >= 2 && value[0] == '"'
                && value[value.Length - 1] == '"') {
                value = value.Substring(1, value.Length - 2).Trim();
            }
            return value.Length == 0 ? null : value;
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
            IReadOnlyDictionary<string, byte[]> streams, out string? reason) {
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
            if (!TryReadRecord(directory, ref position, 0x0003,
                    out int codePageOffset, out int codePageLength)
                || codePageLength != 2
                || !TryReadSizedRecord(directory, ref position, 0x0004, null, true)) {
                reason = "The expanded VBA/dir stream has invalid code-page or project-name records.";
                return false;
            }
            ushort codePage = (ushort)(directory[codePageOffset]
                | directory[codePageOffset + 1] << 8);
            if (!TryReadProjectMetadata(directory, ref position)
                || !TryReadReferences(directory, ref position)
                || !TryReadModules(directory, ref position, codePage,
                    streams)) {
                reason = "The expanded VBA/dir stream contains invalid project, reference, or module records.";
                return false;
            }
            if (!TryReadSizedRecord(directory, ref position, 0x0010, 0,
                    false) || position != directory.Length) {
                reason = "The expanded VBA/dir stream has no valid project terminator.";
                return false;
            }
            reason = null;
            return true;
        }

        private static bool TryReadProjectMetadata(byte[] bytes,
            ref int position) {
            while (true) {
                switch (PeekUInt16(bytes, position)) {
                    case 0x0005:
                        if (!TryReadSizedRecord(bytes, ref position, 0x0005,
                                null, false)
                            || !TryReadSizedRecord(bytes, ref position, 0x0040,
                                null, false)) return false;
                        break;
                    case 0x0006:
                        if (!TryReadSizedRecord(bytes, ref position, 0x0006,
                                null, false)
                            || !TryReadSizedRecord(bytes, ref position, 0x003D,
                                null, false)) return false;
                        break;
                    case 0x0007:
                    case 0x0008:
                        ushort id = PeekUInt16(bytes, position);
                        if (!TryReadSizedRecord(bytes, ref position, id, 4,
                                false)) return false;
                        break;
                    case 0x0009:
                        if (!TryReadUInt16(bytes, ref position, out ushort versionId)
                            || versionId != 0x0009
                            || !TryReadUInt32(bytes, ref position, out uint reserved)
                            || reserved != 4U
                            || !TryReadUInt32(bytes, ref position, out _)
                            || !TryReadUInt16(bytes, ref position, out _)) {
                            return false;
                        }
                        break;
                    case 0x000C:
                        if (!TryReadSizedRecord(bytes, ref position, 0x000C,
                                null, false)
                            || !TryReadSizedRecord(bytes, ref position, 0x003C,
                                null, false)) return false;
                        break;
                    default:
                        return true;
                }
            }
        }

        private static bool TryReadReferences(byte[] bytes, ref int position) {
            while (PeekUInt16(bytes, position) != 0x000F) {
                if (PeekUInt16(bytes, position) == 0x0016
                    && !TryReadReferenceName(bytes, ref position)) return false;
                switch (PeekUInt16(bytes, position)) {
                    case 0x0033:
                        if (!TryReadSizedRecord(bytes, ref position, 0x0033,
                                null, true)) return false;
                        break;
                    case 0x000D:
                        if (!TryReadReferenceRegistered(bytes, ref position)) {
                            return false;
                        }
                        break;
                    case 0x000E:
                        if (!TryReadReferenceProject(bytes, ref position)) {
                            return false;
                        }
                        break;
                    case 0x002F:
                        if (!TryReadReferenceControl(bytes, ref position)) {
                            return false;
                        }
                        break;
                    default:
                        return false;
                }
            }
            return true;
        }

        private static bool TryReadReferenceName(byte[] bytes,
            ref int position) =>
            TryReadSizedRecord(bytes, ref position, 0x0016, null, true)
            && TryReadSizedRecord(bytes, ref position, 0x003E, null, true);

        private static bool TryReadReferenceRegistered(byte[] bytes,
            ref int position) {
            return TryReadSizedRecord(bytes, ref position, 0x000D, null, true)
                && TryReadUInt32(bytes, ref position, out uint reserved1)
                && reserved1 == 0U
                && TryReadUInt16(bytes, ref position, out ushort reserved2)
                && reserved2 == 0U;
        }

        private static bool TryReadReferenceProject(byte[] bytes,
            ref int position) {
            if (!TryReadUInt16(bytes, ref position, out ushort id)
                || id != 0x000E
                || !TryReadLengthPrefixedBytes(bytes, ref position, true)
                || !TryReadLengthPrefixedBytes(bytes, ref position, true)
                || !TryReadUInt32(bytes, ref position, out _)
                || !TryReadUInt16(bytes, ref position, out _)) return false;
            return true;
        }

        private static bool TryReadReferenceControl(byte[] bytes,
            ref int position) {
            if (!TryReadUInt16(bytes, ref position, out ushort id)
                || id != 0x002F
                || !TryReadLengthPrefixedBytes(bytes, ref position, true)
                || !TryReadUInt32(bytes, ref position, out uint reserved1)
                || reserved1 != 0U
                || !TryReadUInt16(bytes, ref position, out ushort reserved2)
                || reserved2 != 0U) return false;
            if (PeekUInt16(bytes, position) == 0x0016
                && !TryReadReferenceName(bytes, ref position)) return false;
            if (!TryReadUInt16(bytes, ref position, out ushort reserved3)
                || reserved3 != 0x0030
                || !TryReadLengthPrefixedBytes(bytes, ref position, true)
                || !TryReadUInt32(bytes, ref position, out uint reserved4)
                || reserved4 != 0U
                || !TryReadUInt16(bytes, ref position, out ushort reserved5)
                || reserved5 != 0U
                || position + 20 > bytes.Length) return false;
            position += 20;
            return true;
        }

        private static bool TryReadModules(byte[] bytes, ref int position,
            ushort codePage, IReadOnlyDictionary<string, byte[]> streams) {
            if (!TryReadRecord(bytes, ref position, 0x000F,
                    out int countOffset, out int countLength)
                || countLength != 2) return false;
            ushort moduleCount = (ushort)(bytes[countOffset]
                | bytes[countOffset + 1] << 8);
            if (moduleCount > MaximumModuleCount) return false;
            if (!TryReadSizedRecord(bytes, ref position, 0x0013, 2, false)) {
                return false;
            }
            var moduleStreams = new Dictionary<string, byte[]>(
                StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, byte[]> pair in streams) {
                if (pair.Key.StartsWith("VBA/", StringComparison.OrdinalIgnoreCase)
                    && !moduleStreams.ContainsKey(pair.Key)) {
                    moduleStreams.Add(pair.Key, pair.Value);
                }
            }
            var validationCache = new Dictionary<string, int?>(
                StringComparer.OrdinalIgnoreCase);
            long cumulativeValidationBytes = 0L;
            for (int index = 0; index < moduleCount; index++) {
                if (!TryReadSizedRecord(bytes, ref position, 0x0019, null, true)) {
                    return false;
                }
                if (PeekUInt16(bytes, position) == 0x0047
                    && !TryReadSizedRecord(bytes, ref position, 0x0047, null,
                        true)) return false;
                if (!TryReadRecord(bytes, ref position, 0x001A,
                        out int streamNameOffset, out int streamNameLength)
                    || streamNameLength == 0
                    || !TryReadRecord(bytes, ref position, 0x0032,
                        out int unicodeStreamNameOffset,
                        out int unicodeStreamNameLength)
                    || unicodeStreamNameLength == 0
                    || unicodeStreamNameLength % 2 != 0
                    || !TryReadSizedRecord(bytes, ref position, 0x001C, null,
                        false)
                    || !TryReadSizedRecord(bytes, ref position, 0x0048, null,
                        false)
                    || !TryReadRecord(bytes, ref position, 0x0031,
                        out int moduleOffsetPosition, out int moduleOffsetLength)
                    || moduleOffsetLength != 4
                    || !TryReadSizedRecord(bytes, ref position, 0x001E, 4,
                        false)
                    || !TryReadSizedRecord(bytes, ref position, 0x002C, 2,
                        false)) return false;
                ushort moduleType = PeekUInt16(bytes, position);
                if ((moduleType != 0x0021 && moduleType != 0x0022)
                    || !TryReadSizedRecord(bytes, ref position, moduleType, 0,
                        false)) return false;
                if (PeekUInt16(bytes, position) == 0x0025
                    && !TryReadSizedRecord(bytes, ref position, 0x0025, 0,
                        false)) return false;
                if (PeekUInt16(bytes, position) == 0x0028
                    && !TryReadSizedRecord(bytes, ref position, 0x0028, 0,
                        false)) return false;
                if (!TryReadSizedRecord(bytes, ref position, 0x002B, 0,
                        false)) return false;

                string? ansiStreamName = TryDecode(bytes, streamNameOffset,
                    streamNameLength, codePage);
                string? unicodeStreamName = TryDecodeUnicode(bytes,
                    unicodeStreamNameOffset, unicodeStreamNameLength);
                string? streamName = !string.IsNullOrWhiteSpace(unicodeStreamName)
                    ? unicodeStreamName
                    : ansiStreamName;
                uint moduleOffset = ReadUInt32(bytes, moduleOffsetPosition);
                string storagePath = "VBA/" + streamName;
                if (string.IsNullOrWhiteSpace(streamName)
                    || !moduleStreams.TryGetValue(storagePath,
                        out byte[]? moduleStream)
                    || moduleOffset >= moduleStream.Length) {
                    return false;
                }
                string cacheKey = storagePath + "\0"
                    + moduleOffset.ToString(CultureInfo.InvariantCulture);
                if (!validationCache.TryGetValue(cacheKey,
                        out int? validationBytes)) {
                    if (!TryValidateCompressedModuleStream(moduleStream,
                            checked((int)moduleOffset),
                            out int measuredBytes)) {
                        validationCache[cacheKey] = null;
                        return false;
                    }
                    validationBytes = measuredBytes;
                    validationCache[cacheKey] = validationBytes;
                    if (validationBytes.Value
                        > MaximumModuleValidationBytes
                            - cumulativeValidationBytes) {
                        return false;
                    }
                    cumulativeValidationBytes += validationBytes.Value;
                } else if (!validationBytes.HasValue) {
                    return false;
                }
            }
            return true;
        }

        private static bool TryValidateCompressedModuleStream(byte[] stream,
            int offset, out int validationBytes) {
            validationBytes = 0;
            int length = stream.Length - offset;
            if (length <= 0) return false;
            var container = new byte[length];
            Buffer.BlockCopy(stream, offset, container, 0, length);
            if (!TryDecompressDirectory(container, out byte[] decompressed,
                    out _)) {
                return false;
            }
            validationBytes = checked(length + decompressed.Length);
            return true;
        }

        private static bool TryReadLengthPrefixedBytes(byte[] bytes,
            ref int position, bool requireData) {
            if (!TryReadUInt32(bytes, ref position, out uint length)
                || length > int.MaxValue || requireData && length == 0
                || position + (long)length > bytes.Length) return false;
            position += (int)length;
            return true;
        }

        private static string? TryDecode(byte[] bytes, int offset, int length,
            ushort codePage) {
            bool ascii = true;
            for (int index = 0; index < length; index++) {
                if (bytes[offset + index] <= 0x7F) continue;
                ascii = false;
                break;
            }
            if (ascii) return Encoding.ASCII.GetString(bytes, offset, length);
            try {
                return Encoding.GetEncoding(codePage)
                    .GetString(bytes, offset, length);
            } catch (Exception exception) when (exception is ArgumentException
                || exception is NotSupportedException) {
                return null;
            }
        }

        private static string? TryDecodeUnicode(byte[] bytes, int offset,
            int length) {
            if (length == 0 || length % 2 != 0) return null;
            try {
                return new UnicodeEncoding(false, false, true)
                    .GetString(bytes, offset, length);
            } catch (DecoderFallbackException) {
                return null;
            }
        }

        private static uint ReadUInt32(byte[] bytes, int position) =>
            (uint)(bytes[position]
                | bytes[position + 1] << 8
                | bytes[position + 2] << 16
                | bytes[position + 3] << 24);

        private static bool TryReadRecord(byte[] bytes, ref int position,
            ushort expectedId, out int dataOffset, out int dataLength) {
            dataOffset = 0;
            dataLength = 0;
            if (!TryReadUInt16(bytes, ref position, out ushort id)
                || id != expectedId
                || !TryReadUInt32(bytes, ref position, out uint size)
                || size > int.MaxValue || position + (long)size > bytes.Length) {
                return false;
            }
            dataOffset = position;
            dataLength = (int)size;
            position += dataLength;
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
