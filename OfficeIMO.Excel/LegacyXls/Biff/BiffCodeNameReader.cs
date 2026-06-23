using OfficeIMO.Excel.LegacyXls.Diagnostics;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffCodeNameReader {
        internal static bool TryRead(BiffRecord record, string? sheetName, List<LegacyXlsImportDiagnostic> diagnostics, out string? codeName) {
            try {
                int offset = 0;
                string value = BiffStringReader.ReadUnicodeString(record.Payload, ref offset).TrimEnd('\0');
                if (value.Length > 31) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-CODENAME-LENGTH-INVALID",
                        "The CodeName record contains a VBA object name longer than 31 characters.",
                        sheetName: sheetName,
                        recordOffset: record.Offset,
                        recordType: record.Type));
                }

                codeName = string.IsNullOrEmpty(value) ? null : value;
                return true;
            } catch (InvalidDataException ex) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-CODENAME-INVALID",
                    $"The CodeName record could not be read. {ex.Message}",
                    sheetName: sheetName,
                    recordOffset: record.Offset,
                    recordType: record.Type));
                codeName = null;
                return false;
            }
        }
    }
}
