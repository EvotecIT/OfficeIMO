using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffProtectedRangeFeatureReader {
        private const ushort FeatRecordType = 0x0868;
        private const ushort IsfProtection = 0x0002;

        internal static bool TryRead(BiffRecord record, out LegacyXlsProtectedRange? protectedRange) {
            protectedRange = null;
            byte[] payload = record.Payload;
            if (record.Type != FeatRecordType || payload.Length < 35) {
                return false;
            }

            ushort rt = BiffRecordReader.ReadUInt16(payload, 0);
            if (rt != FeatRecordType) {
                return false;
            }

            ushort sharedFeatureType = BiffRecordReader.ReadUInt16(payload, 12);
            if (sharedFeatureType != IsfProtection) {
                return false;
            }

            ushort rangeCount = BiffRecordReader.ReadUInt16(payload, 19);
            if (rangeCount == 0 || rangeCount > 432) {
                return false;
            }

            int rangeOffset = 27;
            int featureOffset = checked(rangeOffset + (rangeCount * 8));
            if (featureOffset + 8 > payload.Length) {
                return false;
            }

            List<string> ranges = new(rangeCount);
            for (int i = 0; i < rangeCount; i++) {
                if (!TryReadCellRange(payload, rangeOffset + (i * 8), out string? range)) {
                    return false;
                }

                ranges.Add(range!);
            }

            uint flags = BiffRecordReader.ReadUInt32(payload, featureOffset);
            uint passwordHash = BiffRecordReader.ReadUInt32(payload, featureOffset + 4);
            int stringOffset = featureOffset + 8;
            string title = BiffStringReader.ReadUnicodeString(payload, ref stringOffset);
            if (string.IsNullOrWhiteSpace(title)) {
                return false;
            }

            protectedRange = new LegacyXlsProtectedRange(
                title,
                ranges,
                passwordHash == 0 ? null : ((ushort)passwordHash).ToString("X4"),
                hasSecurityDescriptor: (flags & 0x00000001) != 0);
            return true;
        }

        private static bool TryReadCellRange(byte[] payload, int offset, out string? range) {
            range = null;
            if (offset + 8 > payload.Length) {
                return false;
            }

            ushort firstRow = BiffRecordReader.ReadUInt16(payload, offset);
            ushort lastRow = BiffRecordReader.ReadUInt16(payload, offset + 2);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, offset + 4);
            ushort lastColumn = BiffRecordReader.ReadUInt16(payload, offset + 6);
            if (lastRow < firstRow || lastColumn < firstColumn || firstColumn > 0x00ff || lastColumn > 0x00ff) {
                return false;
            }

            string start = A1.CellReference(firstRow + 1, firstColumn + 1);
            string end = A1.CellReference(lastRow + 1, lastColumn + 1);
            range = start == end ? start : start + ":" + end;
            return true;
        }
    }
}
