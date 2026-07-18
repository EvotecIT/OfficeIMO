namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    internal sealed class LegacyPptRecordTraversalBudget {
        private readonly int _maximumRecordCount;
        private int _recordsTraversed;
        private bool _wasExceeded;

        internal LegacyPptRecordTraversalBudget(int maximumRecordCount) {
            _maximumRecordCount = maximumRecordCount;
        }

        internal int RecordsTraversed => _recordsTraversed;

        internal void ThrowIfExceeded() {
            if (_wasExceeded) {
                throw CreateLimitException();
            }
        }

        internal void Consume() => Consume(1);

        internal void Consume(int recordCount) {
            if (recordCount < 0) {
                throw new ArgumentOutOfRangeException(nameof(recordCount));
            }
            if (_recordsTraversed > _maximumRecordCount - recordCount) {
                _wasExceeded = true;
                throw CreateLimitException();
            }
            _recordsTraversed += recordCount;
        }

        private InvalidDataException CreateLimitException() =>
            new InvalidDataException(
                $"The PowerPoint record count exceeds {_maximumRecordCount}.");
    }

    internal static class LegacyPptRecordReader {
        internal static LegacyPptRecord ReadSingle(byte[] source, int offset, LegacyPptImportOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            return ReadSingle(source, offset, options,
                new LegacyPptRecordTraversalBudget(options.MaxRecordCount));
        }

        internal static LegacyPptRecord ReadSingle(byte[] source, int offset,
            LegacyPptImportOptions options, LegacyPptRecordTraversalBudget budget) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (budget == null) throw new ArgumentNullException(nameof(budget));
            return ReadSingle(source, offset, source.Length, options,
                depth: 0, budget: budget);
        }

        internal static IReadOnlyList<LegacyPptRecord> ReadSequence(byte[] source, int offset, int length,
            LegacyPptImportOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            return ReadSequence(source, offset, length, options,
                new LegacyPptRecordTraversalBudget(options.MaxRecordCount));
        }

        internal static IReadOnlyList<LegacyPptRecord> ReadSequence(byte[] source,
            int offset, int length, LegacyPptImportOptions options,
            LegacyPptRecordTraversalBudget budget) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (budget == null) throw new ArgumentNullException(nameof(budget));
            if (offset < 0 || length < 0 || offset > source.Length - length) {
                throw new InvalidDataException("The PowerPoint record sequence lies outside the stream.");
            }

            return ReadSequence(source, offset, checked(offset + length), options,
                depth: 0, budget: budget);
        }

        private static LegacyPptRecord ReadSingle(byte[] source, int offset, int boundary,
            LegacyPptImportOptions options, int depth,
            LegacyPptRecordTraversalBudget budget) {
            if (depth > options.MaxRecordDepth) {
                throw new InvalidDataException($"The PowerPoint record nesting depth exceeds {options.MaxRecordDepth}.");
            }
            budget.Consume();
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
                children = ReadSequence(source, payloadOffset, endOffset, options,
                        checked(depth + 1), budget)
                    .ToList();
            }
            return new LegacyPptRecord(source, offset, version, instance, type, payloadOffset, payloadLength, children);
        }

        private static IReadOnlyList<LegacyPptRecord> ReadSequence(byte[] source, int offset, int endOffset,
            LegacyPptImportOptions options, int depth,
            LegacyPptRecordTraversalBudget budget) {
            var records = new List<LegacyPptRecord>();
            int position = offset;
            while (position < endOffset) {
                if (endOffset - position < 8) {
                    throw new InvalidDataException(
                        $"The PowerPoint record sequence has {endOffset - position} trailing byte(s) at 0x{position:X}.");
                }
                LegacyPptRecord record = ReadSingle(source, position, endOffset,
                    options, depth, budget);
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
