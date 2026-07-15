using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    internal static class LegacyPptDocumentAtomReader {
        internal const int PayloadLength = 40;

        internal static LegacyPptDocumentSettings? Read(LegacyPptRecord? record) {
            if (record == null || record.PayloadLength < PayloadLength) return null;
            return new LegacyPptDocumentSettings(
                record.ReadInt32(0), record.ReadInt32(4),
                record.ReadInt32(8), record.ReadInt32(12),
                record.ReadInt32(16), record.ReadInt32(20),
                record.ReadUInt32(24), record.ReadUInt32(28),
                record.ReadUInt16(32), record.ReadUInt16(34),
                record.ReadByte(36) != 0, record.ReadByte(37) != 0,
                record.ReadByte(38) != 0, record.ReadByte(39) != 0);
        }
    }
}
