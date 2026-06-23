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

            TryReadObjectCommonData(record, out ushort? objectType, out ushort? objectId);
            records.Add(new LegacyXlsDrawingRecord(
                GetKind(record.Type),
                BiffUnsupportedRecordDiagnostics.GetBiffRecordName(record.Type),
                sheetName,
                record.Offset,
                record.Type,
                record.Payload.Length,
                objectType,
                objectId));
            return true;
        }

        private static bool TryReadObjectCommonData(BiffRecord record, out ushort? objectType, out ushort? objectId) {
            objectType = null;
            objectId = null;
            byte[] payload = record.Payload;
            if (record.Type != (ushort)BiffRecordType.Obj || payload.Length < 8) {
                return false;
            }

            ushort subRecordType = BiffRecordReader.ReadUInt16(payload, 0);
            ushort subRecordLength = BiffRecordReader.ReadUInt16(payload, 2);
            if (subRecordType != 0x0015 || subRecordLength < 4 || payload.Length < 8) {
                return false;
            }

            objectType = BiffRecordReader.ReadUInt16(payload, 4);
            objectId = BiffRecordReader.ReadUInt16(payload, 6);
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
