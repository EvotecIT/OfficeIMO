namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Provides bounded little-endian reads over a binary PowerPoint text-property atom.</summary>
    internal sealed class LegacyPptTextPropertyCursor {
        private readonly LegacyPptRecord _record;
        private readonly string _description;

        internal LegacyPptTextPropertyCursor(LegacyPptRecord record, string description) {
            _record = record ?? throw new ArgumentNullException(nameof(record));
            _description = string.IsNullOrWhiteSpace(description) ? "Text property atom" : description;
        }

        internal int Offset { get; private set; }

        internal int Remaining => _record.PayloadLength - Offset;

        internal bool IsAtEnd => Offset == _record.PayloadLength;

        internal byte ReadByte() {
            byte value = _record.ReadByte(Offset);
            Offset = checked(Offset + 1);
            return value;
        }

        internal ushort ReadUInt16() {
            ushort value = _record.ReadUInt16(Offset);
            Offset = checked(Offset + 2);
            return value;
        }

        internal short ReadInt16() => unchecked((short)ReadUInt16());

        internal uint ReadUInt32() {
            uint value = _record.ReadUInt32(Offset);
            Offset = checked(Offset + 4);
            return value;
        }

        internal void Skip(int count) {
            if (count < 0 || Offset > _record.PayloadLength - count) {
                throw new InvalidDataException($"{_description} is truncated.");
            }
            Offset = checked(Offset + count);
        }
    }
}
