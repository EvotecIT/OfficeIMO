using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffDrawingMetadataReader {
        internal static bool TryRead(
            BiffRecord record,
            string? sheetName,
            List<LegacyXlsDrawingRecord> records) {
            if (!BiffUnsupportedRecordDiagnostics.IsDrawingRecord(record.Type)) {
                return false;
            }

            records.Add(new LegacyXlsDrawingRecord(
                GetKind(record.Type),
                BiffUnsupportedRecordDiagnostics.GetBiffRecordName(record.Type),
                sheetName,
                record.Offset,
                record.Type,
                record.Payload.Length));
            return true;
        }

        private static LegacyXlsDrawingRecordKind GetKind(ushort type) {
            if (type == (ushort)BiffRecordType.DrawingGroup) {
                return LegacyXlsDrawingRecordKind.DrawingGroup;
            }

            if (type == (ushort)BiffRecordType.Drawing) {
                return LegacyXlsDrawingRecordKind.Drawing;
            }

            if (type == (ushort)BiffRecordType.Obj) {
                return LegacyXlsDrawingRecordKind.Object;
            }

            if (type == (ushort)BiffRecordType.Txo) {
                return LegacyXlsDrawingRecordKind.TextObject;
            }

            return LegacyXlsDrawingRecordKind.PreserveOnly;
        }
    }
}
