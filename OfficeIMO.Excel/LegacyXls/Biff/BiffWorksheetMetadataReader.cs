using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffWorksheetMetadataReader {
        internal static bool ReadGridSet(BiffRecord record, string sheetName, List<LegacyXlsImportDiagnostic> diagnostics) {
            if (record.Payload.Length < 2) {
                throw new InvalidDataException("The GRIDSET record is shorter than expected.");
            }

            ushort value = BiffRecordReader.ReadUInt16(record.Payload, 0);
            if (value != 0 && value != 1) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-GRIDSET-VALUE-UNEXPECTED",
                    $"The GRIDSET record contains unexpected value {value}.",
                    sheetName: sheetName,
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }

            return value != 0;
        }

        internal static (byte RowLevel, byte ColumnLevel) ReadGuts(byte[] payload) {
            if (payload.Length < 8) {
                throw new InvalidDataException("The GUTS record is shorter than expected.");
            }

            return (
                ToOutlineLevel(BiffRecordReader.ReadUInt16(payload, 4), "row"),
                ToOutlineLevel(BiffRecordReader.ReadUInt16(payload, 6), "column"));
        }

        internal static LegacyXlsWorksheetIndex ReadIndex(byte[] payload, bool useBiff5Layout) {
            if (payload.Length < 16) {
                throw new InvalidDataException("The INDEX record is shorter than expected.");
            }

            if (useBiff5Layout && payload.Length == 16) {
                ushort biff5FirstRow = BiffRecordReader.ReadUInt16(payload, 4);
                ushort biff5RowAfterLast = BiffRecordReader.ReadUInt16(payload, 6);
                if (biff5RowAfterLast < biff5FirstRow) {
                    throw new InvalidDataException("The INDEX record contains an invalid row range.");
                }

                return new LegacyXlsWorksheetIndex(
                    biff5FirstRow + 1,
                    biff5RowAfterLast + 1,
                    BiffRecordReader.ReadUInt32(payload, 8),
                    dbCellBlockCount: 2);
            }

            uint firstRow = BiffRecordReader.ReadUInt32(payload, 4);
            uint rowAfterLast = BiffRecordReader.ReadUInt32(payload, 8);
            if (rowAfterLast < firstRow) {
                throw new InvalidDataException("The INDEX record contains an invalid row range.");
            }

            int dbCellBlockCount = checked((payload.Length - 16) / 4);
            if (payload.Length != 16 + (dbCellBlockCount * 4)) {
                throw new InvalidDataException("The INDEX record contains a partial DBCell offset.");
            }

            return new LegacyXlsWorksheetIndex(
                checked((int)firstRow + 1),
                checked((int)rowAfterLast + 1),
                BiffRecordReader.ReadUInt32(payload, 12),
                dbCellBlockCount);
        }

        internal static LegacyXlsSelection ReadSelection(byte[] payload) {
            if (payload.Length < 9) {
                throw new InvalidDataException("The SELECTION record is shorter than expected.");
            }

            byte pane = payload[0];
            int activeRow = BiffRecordReader.ReadUInt16(payload, 1) + 1;
            int activeColumn = BiffRecordReader.ReadUInt16(payload, 3) + 1;
            ushort activeRangeIndex = BiffRecordReader.ReadUInt16(payload, 5);
            ushort rangeCount = BiffRecordReader.ReadUInt16(payload, 7);
            int expectedLength = checked(9 + (rangeCount * 6));
            if (expectedLength > payload.Length) {
                throw new InvalidDataException("The SELECTION record ended before all selected ranges could be read.");
            }

            var ranges = new List<LegacyXlsSelectedRange>(rangeCount);
            for (int i = 0; i < rangeCount; i++) {
                int rangeOffset = 9 + (i * 6);
                int firstRow = BiffRecordReader.ReadUInt16(payload, rangeOffset) + 1;
                int lastRow = BiffRecordReader.ReadUInt16(payload, rangeOffset + 2) + 1;
                int firstColumn = payload[rangeOffset + 4] + 1;
                int lastColumn = payload[rangeOffset + 5] + 1;
                if (lastRow < firstRow || lastColumn < firstColumn) {
                    throw new InvalidDataException("The SELECTION record contains an invalid selected range.");
                }

                ranges.Add(new LegacyXlsSelectedRange(firstRow, firstColumn, lastRow, lastColumn));
            }

            return new LegacyXlsSelection(pane, activeRow, activeColumn, activeRangeIndex, ranges);
        }

        internal static ushort ReadWsBool(byte[] payload) {
            if (payload.Length < 2) {
                throw new InvalidDataException("The WSBOOL record is shorter than expected.");
            }

            return BiffRecordReader.ReadUInt16(payload, 0);
        }

        private static byte ToOutlineLevel(ushort rawLevel, string axis) {
            if (rawLevel == 0) {
                return 0;
            }

            if (rawLevel < 2 || rawLevel > 8) {
                throw new InvalidDataException($"The GUTS record contains an invalid {axis} outline level.");
            }

            return checked((byte)(rawLevel - 1));
        }
    }
}
