using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffIgnoredErrorsFeatureReader {
        private const ushort FeatHdrRecordType = 0x0867;
        private const ushort FeatRecordType = 0x0868;
        private const ushort IsfIgnoredErrors = 0x0003;
        private const uint IgnoredErrorsDataSize = 4;

        internal static bool TryReadHeader(BiffRecord record) {
            byte[] payload = record.Payload;
            if (record.Type != FeatHdrRecordType || payload.Length < 19) {
                return false;
            }

            ushort rt = BiffRecordReader.ReadUInt16(payload, 0);
            ushort sharedFeatureType = BiffRecordReader.ReadUInt16(payload, 12);
            uint headerDataSize = BiffRecordReader.ReadUInt32(payload, 15);
            return rt == FeatHdrRecordType
                && sharedFeatureType == IsfIgnoredErrors
                && payload[14] == 1
                && headerDataSize == 0;
        }

        internal static bool TryRead(BiffRecord record, out LegacyXlsIgnoredError? ignoredError) {
            ignoredError = null;
            byte[] payload = record.Payload;
            if (record.Type != FeatRecordType || payload.Length < 31) {
                return false;
            }

            ushort rt = BiffRecordReader.ReadUInt16(payload, 0);
            if (rt != FeatRecordType) {
                return false;
            }

            ushort sharedFeatureType = BiffRecordReader.ReadUInt16(payload, 12);
            if (sharedFeatureType != IsfIgnoredErrors) {
                return false;
            }

            ushort rangeCount = BiffRecordReader.ReadUInt16(payload, 19);
            uint featureDataSize = BiffRecordReader.ReadUInt32(payload, 21);
            if (rangeCount == 0 || rangeCount > 1027 || featureDataSize != IgnoredErrorsDataSize) {
                return false;
            }

            int rangeOffset = 27;
            int featureOffset = checked(rangeOffset + (rangeCount * 8));
            if (featureOffset + 4 > payload.Length) {
                return false;
            }

            var references = new List<string>(rangeCount);
            for (int i = 0; i < rangeCount; i++) {
                if (!TryReadCellRange(payload, rangeOffset + (i * 8), out string? reference)) {
                    return false;
                }

                references.Add(reference!);
            }

            uint flags = BiffRecordReader.ReadUInt32(payload, featureOffset);
            if ((flags & 0xffffff00U) != 0) {
                return false;
            }

            ignoredError = new LegacyXlsIgnoredError(
                references,
                evaluationError: IsBitSet(flags, 0),
                emptyCellReference: IsBitSet(flags, 1),
                numberStoredAsText: IsBitSet(flags, 2),
                formulaRange: IsBitSet(flags, 3),
                formula: IsBitSet(flags, 4),
                twoDigitTextYear: IsBitSet(flags, 5),
                unlockedFormula: IsBitSet(flags, 6),
                listDataValidation: IsBitSet(flags, 7));
            return true;
        }

        private static bool TryReadCellRange(byte[] payload, int offset, out string? reference) {
            reference = null;
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
            reference = start == end ? start : start + ":" + end;
            return true;
        }

        private static bool IsBitSet(uint value, int bit) {
            return (value & (1U << bit)) != 0;
        }
    }
}
