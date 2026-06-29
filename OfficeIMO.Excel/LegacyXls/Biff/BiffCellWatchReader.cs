using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffCellWatchReader {
        private const ushort CellWatchRecordType = 0x086C;

        internal static bool TryRead(BiffRecord record, out LegacyXlsCellWatch? cellWatch) {
            cellWatch = null;
            byte[] payload = record.Payload;
            if (record.Type != CellWatchRecordType || payload.Length != 16) {
                return false;
            }

            ushort rt = BiffRecordReader.ReadUInt16(payload, 0);
            ushort grbitFrt = BiffRecordReader.ReadUInt16(payload, 2);
            if (rt != CellWatchRecordType || (grbitFrt & 0x0001) == 0) {
                return false;
            }

            ushort firstRow = BiffRecordReader.ReadUInt16(payload, 4);
            ushort lastRow = BiffRecordReader.ReadUInt16(payload, 6);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, 8);
            ushort lastColumn = BiffRecordReader.ReadUInt16(payload, 10);
            uint reserved = BiffRecordReader.ReadUInt32(payload, 12);
            if (reserved != 0 || firstRow != lastRow || firstColumn != lastColumn || firstColumn > 0x00ff) {
                return false;
            }

            int row = firstRow + 1;
            int column = firstColumn + 1;
            cellWatch = new LegacyXlsCellWatch(A1.CellReference(row, column), row, column);
            return true;
        }
    }
}
