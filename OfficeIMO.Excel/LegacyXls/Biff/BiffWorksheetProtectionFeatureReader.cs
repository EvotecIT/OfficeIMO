using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffWorksheetProtectionFeatureReader {
        private const ushort FeatHdrRecordType = 0x0867;
        private const ushort IsfProtection = 0x0002;
        private const uint FeatHdrProtectionDataSize = 0xffffffff;

        internal static bool TryRead(BiffRecord record, out LegacyXlsWorksheetProtectionPermissions? permissions) {
            permissions = null;
            byte[] payload = record.Payload;
            if (record.Type != FeatHdrRecordType || payload.Length < 23) {
                return false;
            }

            ushort rt = BiffRecordReader.ReadUInt16(payload, 0);
            if (rt != FeatHdrRecordType) {
                return false;
            }

            ushort sharedFeatureType = BiffRecordReader.ReadUInt16(payload, 12);
            uint headerDataSize = BiffRecordReader.ReadUInt32(payload, 15);
            if (sharedFeatureType != IsfProtection || headerDataSize != FeatHdrProtectionDataSize) {
                return false;
            }

            uint flags = BiffRecordReader.ReadUInt32(payload, 19);
            permissions = new LegacyXlsWorksheetProtectionPermissions(
                allowEditObjects: IsBitSet(flags, 2),
                allowEditScenarios: IsBitSet(flags, 3),
                allowFormatCells: IsBitSet(flags, 4),
                allowFormatColumns: IsBitSet(flags, 5),
                allowFormatRows: IsBitSet(flags, 6),
                allowInsertColumns: IsBitSet(flags, 7),
                allowInsertRows: IsBitSet(flags, 8),
                allowInsertHyperlinks: IsBitSet(flags, 9),
                allowDeleteColumns: IsBitSet(flags, 10),
                allowDeleteRows: IsBitSet(flags, 11),
                allowSelectLockedCells: IsBitSet(flags, 12),
                allowSort: IsBitSet(flags, 13),
                allowAutoFilter: IsBitSet(flags, 14),
                allowPivotTables: IsBitSet(flags, 15),
                allowSelectUnlockedCells: IsBitSet(flags, 16));
            return true;
        }

        private static bool IsBitSet(uint value, int bit) {
            return (value & (1u << bit)) != 0;
        }
    }
}
