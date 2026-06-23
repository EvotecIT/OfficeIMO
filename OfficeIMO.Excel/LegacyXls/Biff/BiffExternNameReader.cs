using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Reads ExternName records attached to a preceding SupBook supporting link.
    /// </summary>
    internal static class BiffExternNameReader {
        internal static bool TryRead(
            BiffRecord record,
            LegacyXlsExternalReferenceKind referenceKind,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out LegacyXlsExternalName? externalName) {
            externalName = null;
            try {
                if (record.Payload.Length < 7) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-EXTERNNAME-SHORT",
                        "An ExternName record was too short to parse.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    return false;
                }

                ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 0);
                bool builtIn = (flags & 0x0001) != 0;
                int offset = 2;
                ushort oneBasedSheetIndex = 0;
                if (referenceKind == LegacyXlsExternalReferenceKind.AddIn) {
                    offset += 4; // skip AddinUdf.reserved
                } else {
                    oneBasedSheetIndex = BiffRecordReader.ReadUInt16(record.Payload, offset);
                    offset += 4; // skip ixals and reserved
                }

                string rawName = BiffStringReader.ReadShortUnicodeString(record.Payload, ref offset);
                string name = builtIn ? GetBuiltInName(rawName) ?? string.Empty : rawName;
                if (string.IsNullOrWhiteSpace(name)) {
                    return false;
                }

                int? localSheetIndex = oneBasedSheetIndex == 0 ? null : oneBasedSheetIndex - 1;
                externalName = new LegacyXlsExternalName(name, localSheetIndex, builtIn);
                return true;
            } catch (Exception ex) when (ex is InvalidDataException || ex is OverflowException) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-EXTERNNAME-INVALID",
                    $"An ExternName record could not be parsed. {ex.Message}",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return false;
            }
        }

        private static string? GetBuiltInName(string rawName) {
            if (rawName.Length != 1) {
                return null;
            }

            switch (rawName[0]) {
                case (char)0x06:
                    return "_xlnm.Print_Area";
                case (char)0x07:
                    return "_xlnm.Print_Titles";
                case (char)0x0d:
                    return "_FilterDatabase";
                default:
                    return null;
            }
        }
    }
}
