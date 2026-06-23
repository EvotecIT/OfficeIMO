using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffPivotTableMetadataReader {
        internal static bool TryRead(
            BiffRecord record,
            string? sheetName,
            List<LegacyXlsPivotTableRecord> records,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (!BiffUnsupportedRecordDiagnostics.IsPivotTableRecord(record.Type)) {
                return false;
            }

            LegacyXlsPivotTableRecord pivotRecord = CreateRecord(record, sheetName);
            records.Add(pivotRecord);
            try {
                switch (record.Type) {
                    case 0x00C1:
                        ReadDataItem(record, pivotRecord);
                        break;
                    case 0x00D7:
                        ReadGroupingRange(record, pivotRecord);
                        break;
                    case 0x00FF:
                        ReadExtendedPivotField(record, pivotRecord);
                        break;
                }
            } catch (InvalidDataException) {
                // PivotTable import is currently preserve-only. Keep the typed record node,
                // but avoid turning partial metadata decode into an import warning.
            }

            return true;
        }

        private static LegacyXlsPivotTableRecord CreateRecord(BiffRecord record, string? sheetName) {
            LegacyXlsPivotTableRecordKind kind;
            switch (record.Type) {
                case 0x00C1:
                    kind = LegacyXlsPivotTableRecordKind.DataItem;
                    break;
                case 0x00D7:
                    kind = LegacyXlsPivotTableRecordKind.GroupingRange;
                    break;
                case 0x00FF:
                    kind = LegacyXlsPivotTableRecordKind.ExtendedPivotField;
                    break;
                default:
                    kind = LegacyXlsPivotTableRecordKind.PreserveOnly;
                    break;
            }

            return new LegacyXlsPivotTableRecord(
                kind,
                BiffUnsupportedRecordDiagnostics.GetBiffRecordName(record.Type),
                sheetName,
                record.Offset,
                record.Type,
                record.Payload.Length);
        }

        private static void ReadDataItem(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 14) {
                throw new InvalidDataException("The SXDI payload is shorter than the fixed data item header.");
            }

            short dataItemFieldIndex = BiffRecordReader.ReadInt16(payload, 0);
            short aggregationFunction = BiffRecordReader.ReadInt16(payload, 2);
            short displayCalculation = BiffRecordReader.ReadInt16(payload, 4);
            short displayCalculationFieldIndex = BiffRecordReader.ReadInt16(payload, 6);
            short displayCalculationItemIndex = BiffRecordReader.ReadInt16(payload, 8);
            ushort numberFormatId = BiffRecordReader.ReadUInt16(payload, 10);
            ushort nameLength = BiffRecordReader.ReadUInt16(payload, 12);
            string? name = null;
            if (nameLength != 0xFFFF) {
                if (nameLength == 0) {
                    throw new InvalidDataException("The SXDI custom name length is zero.");
                }

                int offset = 14;
                name = BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, nameLength);
            }

            pivotRecord.SetDataItem(
                dataItemFieldIndex,
                aggregationFunction,
                displayCalculation,
                displayCalculationFieldIndex,
                displayCalculationItemIndex,
                numberFormatId,
                name);
        }

        private static void ReadGroupingRange(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 2) {
                throw new InvalidDataException("The SXRng payload is shorter than the grouping flags.");
            }

            ushort flags = BiffRecordReader.ReadUInt16(payload, 0);
            bool autoStart = (flags & 0x0001) != 0;
            bool autoEnd = (flags & 0x0002) != 0;
            int groupingValue = (flags >> 2) & 0x0007;
            pivotRecord.SetGroupingRange(autoStart, autoEnd, (LegacyXlsPivotGroupingKind)groupingValue);
        }

        private static void ReadExtendedPivotField(BiffRecord record, LegacyXlsPivotTableRecord pivotRecord) {
            byte[] payload = record.Payload;
            if (payload.Length < 2) {
                throw new InvalidDataException("The SXVDEx payload is shorter than the extended pivot field flags.");
            }

            ushort flags = BiffRecordReader.ReadUInt16(payload, 0);
            pivotRecord.SetExtendedPivotField(
                showAllItems: (flags & 0x0001) != 0,
                canDragToRow: (flags & 0x0002) != 0,
                canDragToColumn: (flags & 0x0004) != 0,
                canDragToPage: (flags & 0x0008) != 0,
                canDragToHide: (flags & 0x0010) != 0,
                preventDragToData: (flags & 0x0020) != 0,
                serverBased: (flags & 0x0080) != 0);
        }
    }
}
