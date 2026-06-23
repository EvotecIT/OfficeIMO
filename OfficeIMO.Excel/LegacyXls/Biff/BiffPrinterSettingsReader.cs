using OfficeIMO.Excel.LegacyXls.Diagnostics;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffPrinterSettingsReader {
        internal static void Validate(BiffRecord record, string? sheetName, List<LegacyXlsImportDiagnostic> diagnostics) {
            if (record.Payload.Length < 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-PLS-SHORT",
                    "The Pls printer settings record is shorter than expected.",
                    sheetName: sheetName,
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return;
            }

            ushort reserved = BiffRecordReader.ReadUInt16(record.Payload, 0);
            if (reserved != 0) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-PLS-RESERVED-VALUE-UNEXPECTED",
                    $"The Pls printer settings record contains unexpected reserved value {reserved}.",
                    sheetName: sheetName,
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }
        }
    }
}
