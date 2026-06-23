using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffStyleReader {
        internal static bool TryRead(BiffRecord record, LegacyXlsWorkbook workbook, List<LegacyXlsImportDiagnostic> diagnostics) {
            if (record.Type != (ushort)BiffRecordType.Style) {
                return false;
            }

            if (TryReadStyle(record, diagnostics, out LegacyXlsCellStyle? style)) {
                workbook.AddCellStyle(style!);
            }

            return true;
        }

        private static bool TryReadStyle(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out LegacyXlsCellStyle? style) {
            if (record.Payload.Length < 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-STYLE-SHORT",
                    "The Style record is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                style = null;
                return false;
            }

            ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 0);
            ushort styleFormatIndex = (ushort)(flags & 0x0fff);
            bool isBuiltIn = (flags & 0x8000) != 0;
            if (isBuiltIn) {
                if (record.Payload.Length < 4) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-STYLE-BUILTIN-SHORT",
                        "The built-in Style record is missing BuiltInStyle data.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    style = null;
                    return false;
                }

                style = new LegacyXlsCellStyle(
                    styleFormatIndex,
                    isBuiltIn: true,
                    builtInStyleId: record.Payload[2],
                    outlineLevel: record.Payload[3],
                    name: null,
                    record.Offset,
                    record.Type);
                return true;
            }

            if (record.Payload.Length == 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-STYLE-NAME-MISSING",
                    "The custom Style record is missing its style name.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                style = null;
                return false;
            }

            try {
                int offset = 2;
                string name = BiffStringReader.ReadUnicodeString(record.Payload, ref offset);
                style = new LegacyXlsCellStyle(
                    styleFormatIndex,
                    isBuiltIn: false,
                    builtInStyleId: null,
                    outlineLevel: null,
                    name,
                    record.Offset,
                    record.Type);
                return true;
            } catch (InvalidDataException ex) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-STYLE-NAME-INVALID",
                    $"The custom Style record name could not be read. {ex.Message}",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                style = null;
                return false;
            }
        }
    }
}
