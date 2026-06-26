using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Reads SupBook supporting-link records from BIFF workbook globals.
    /// </summary>
    internal static class BiffSupBookReader {
        internal static bool TryRead(
            BiffRecord record,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out LegacyXlsExternalReference? reference) {
            reference = null;
            try {
                if (record.Payload.Length < 4) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-SUPBOOK-SHORT",
                        "A SupBook record was too short to parse.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    return false;
                }

                ushort sheetCount = BiffRecordReader.ReadUInt16(record.Payload, 0);
                ushort characterCount = BiffRecordReader.ReadUInt16(record.Payload, 2);
                reference = ReadReference(record.Payload, sheetCount, characterCount);
                return true;
            } catch (Exception ex) when (ex is InvalidDataException || ex is OverflowException) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-SUPBOOK-INVALID",
                    $"A SupBook record could not be parsed. {ex.Message}",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return false;
            }
        }

        private static LegacyXlsExternalReference ReadReference(byte[] payload, ushort sheetCount, ushort characterCount) {
            if (characterCount == 0x0401) {
                return new LegacyXlsExternalReference(LegacyXlsExternalReferenceKind.Self, null, null, sheetCount);
            }

            if (characterCount == 0x3A01) {
                return new LegacyXlsExternalReference(LegacyXlsExternalReferenceKind.AddIn, null, null, sheetCount);
            }

            if (characterCount == 0 || characterCount > 0x00ff) {
                return new LegacyXlsExternalReference(LegacyXlsExternalReferenceKind.Unknown, null, null, sheetCount);
            }

            int offset = 4;
            string target = BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, characterCount);
            if (target.Length == 1 && target[0] == '\0') {
                return new LegacyXlsExternalReference(LegacyXlsExternalReferenceKind.SameSheet, null, null, sheetCount);
            }

            var sheetNames = new List<string>();
            for (int i = 0; i < sheetCount && offset < payload.Length; i++) {
                sheetNames.Add(BiffStringReader.ReadUnicodeString(payload, ref offset));
            }

            if (target.Length == 1 && target[0] == ' ') {
                return new LegacyXlsExternalReference(LegacyXlsExternalReferenceKind.Unused, null, sheetNames, sheetCount);
            }

            LegacyXlsExternalReferenceKind kind = sheetCount > 0
                ? LegacyXlsExternalReferenceKind.ExternalWorkbook
                : LegacyXlsExternalReferenceKind.DdeOrOle;
            return new LegacyXlsExternalReference(kind, target, sheetNames, sheetCount);
        }
    }
}
