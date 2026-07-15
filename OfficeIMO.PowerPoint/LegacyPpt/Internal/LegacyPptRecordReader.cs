namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    internal static class LegacyPptRecordReader {
        internal static LegacyPptRecord ReadSingle(byte[] source, int offset, LegacyPptImportOptions options) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));
            int recordCount = 0;
            return ReadSingle(source, offset, source.Length, options, depth: 0, ref recordCount);
        }

        internal static IReadOnlyList<LegacyPptRecord> ReadSequence(byte[] source, int offset, int length,
            LegacyPptImportOptions options) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (offset < 0 || length < 0 || offset > source.Length - length) {
                throw new InvalidDataException("The PowerPoint record sequence lies outside the stream.");
            }

            int recordCount = 0;
            return ReadSequence(source, offset, checked(offset + length), options, depth: 0, ref recordCount);
        }

        private static LegacyPptRecord ReadSingle(byte[] source, int offset, int boundary,
            LegacyPptImportOptions options, int depth, ref int recordCount) {
            if (depth > options.MaxRecordDepth) {
                throw new InvalidDataException($"The PowerPoint record nesting depth exceeds {options.MaxRecordDepth}.");
            }
            if (++recordCount > options.MaxRecordCount) {
                throw new InvalidDataException($"The PowerPoint record count exceeds {options.MaxRecordCount}.");
            }
            if (offset < 0 || offset > boundary - 8 || boundary > source.Length) {
                throw new InvalidDataException($"A PowerPoint record header at 0x{offset:X} is truncated.");
            }

            ushort versionAndInstance = ReadUInt16(source, offset);
            byte version = unchecked((byte)(versionAndInstance & 0x000F));
            ushort instance = unchecked((ushort)(versionAndInstance >> 4));
            ushort type = ReadUInt16(source, offset + 2);
            uint declaredLength = ReadUInt32(source, offset + 4);
            if (declaredLength > int.MaxValue) {
                throw new InvalidDataException($"PowerPoint record 0x{type:X4} at 0x{offset:X} is too large.");
            }

            int payloadLength = unchecked((int)declaredLength);
            int payloadOffset = checked(offset + 8);
            int endOffset;
            try {
                endOffset = checked(payloadOffset + payloadLength);
            } catch (OverflowException exception) {
                throw new InvalidDataException($"PowerPoint record 0x{type:X4} at 0x{offset:X} overflows the stream.", exception);
            }
            if (endOffset > boundary) {
                throw new InvalidDataException(
                    $"PowerPoint record 0x{type:X4} at 0x{offset:X} declares {payloadLength} payload bytes beyond its container.");
            }

            List<LegacyPptRecord>? children = null;
            if (version == 0x0F && payloadLength > 0) {
                children = ReadSequence(source, payloadOffset, endOffset, options, checked(depth + 1), ref recordCount)
                    .ToList();
            }
            return new LegacyPptRecord(source, offset, version, instance, type, payloadOffset, payloadLength, children);
        }

        private static IReadOnlyList<LegacyPptRecord> ReadSequence(byte[] source, int offset, int endOffset,
            LegacyPptImportOptions options, int depth, ref int recordCount) {
            var records = new List<LegacyPptRecord>();
            int position = offset;
            while (position < endOffset) {
                if (endOffset - position < 8) {
                    throw new InvalidDataException(
                        $"The PowerPoint record sequence has {endOffset - position} trailing byte(s) at 0x{position:X}.");
                }
                LegacyPptRecord record = ReadSingle(source, position, endOffset, options, depth, ref recordCount);
                records.Add(record);
                position = record.EndOffset;
            }
            return records;
        }

        private static ushort ReadUInt16(byte[] bytes, int offset) => unchecked((ushort)(bytes[offset]
            | (bytes[offset + 1] << 8)));

        private static uint ReadUInt32(byte[] bytes, int offset) => unchecked((uint)(bytes[offset]
            | (bytes[offset + 1] << 8)
            | (bytes[offset + 2] << 16)
            | (bytes[offset + 3] << 24)));
    }
}
