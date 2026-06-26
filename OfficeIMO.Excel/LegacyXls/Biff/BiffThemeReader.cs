using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffThemeReader {
        private const int FrtHeaderSize = 12;
        private const int ThemeVersionOffset = FrtHeaderSize;
        private const int ThemePayloadMinimumLength = FrtHeaderSize + 4;
        private const uint CustomThemeVersion = 0;
        private const uint DefaultThemeVersion = 124226;

        internal static bool TryRead(
            BiffRecord record,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out LegacyXlsThemeRecord? themeRecord) {
            themeRecord = null;
            if (record.Type != (ushort)BiffRecordType.Theme) {
                return false;
            }

            if (record.Payload.Length < ThemePayloadMinimumLength) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-THEME-SHORT",
                    "The Theme record is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return true;
            }

            ushort headerRecordType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            if (headerRecordType != (ushort)BiffRecordType.Theme) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-THEME-HEADER-UNEXPECTED",
                    $"The Theme future record header declares record type 0x{headerRecordType:X4}.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }

            uint themeVersion = BiffRecordReader.ReadUInt32(record.Payload, ThemeVersionOffset);
            string themeVersionName = GetThemeVersionName(themeVersion);

            byte[] themeBytes = Array.Empty<byte>();
            if (record.Payload.Length > ThemePayloadMinimumLength) {
                themeBytes = new byte[record.Payload.Length - ThemePayloadMinimumLength];
                Buffer.BlockCopy(record.Payload, ThemePayloadMinimumLength, themeBytes, 0, themeBytes.Length);
            }

            themeRecord = new LegacyXlsThemeRecord(
                record.Offset,
                record.Type,
                record.Payload.Length,
                themeVersion,
                themeVersionName,
                themeBytes);
            return true;
        }

        private static string GetThemeVersionName(uint themeVersion) {
            switch (themeVersion) {
                case CustomThemeVersion:
                    return "Custom";
                case DefaultThemeVersion:
                    return "Default";
                default:
                    return $"Version:{themeVersion}";
            }
        }
    }
}
