using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Validates BIFF BOF records before version-specific record layouts are interpreted.
    /// </summary>
    internal static class LegacyBiffVersionValidator {
        internal const ushort Biff5Version = 0x0500;
        internal const ushort Biff8Version = 0x0600;
        private const ushort WorkbookGlobalsSubstream = 0x0005;

        internal static bool ValidateWorkbookGlobals(
            IReadOnlyList<BiffRecord> records,
            LegacyXlsWorkbook workbook) {
            if (records.Count == 0) {
                return false;
            }

            BiffRecord first = records[0];
            if (first.Type != (ushort)BiffRecordType.Bof) {
                workbook.MutableDiagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-BIFF-BOF-MISSING",
                    "The BIFF workbook stream does not start with a BOF record.",
                    recordOffset: first.Offset,
                    recordType: first.Type));
                return false;
            }

            return ValidateBofPayload(first.Payload, first.Offset, WorkbookGlobalsSubstream, sheetName: null, workbook.MutableUnsupportedFeatures, workbook.MutablePreservedFeatureRecords, workbook.MutableDiagnostics);
        }

        internal static bool ValidateWorksheetBof(
            byte[] payload,
            int offset,
            string sheetName,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            return ValidateBofPayload(payload, offset, expectedSubstreamType: null, sheetName, unsupportedFeatures, preservedFeatureRecords, diagnostics);
        }

        private static bool ValidateBofPayload(
            byte[] payload,
            int offset,
            ushort? expectedSubstreamType,
            string? sheetName,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (payload.Length < 4) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-BIFF-BOF-SHORT",
                    "A BIFF BOF record was too short to identify the workbook version and substream type.",
                    sheetName: sheetName,
                    recordOffset: offset,
                    recordType: (ushort)BiffRecordType.Bof));
                return false;
            }

            ushort version = BiffRecordReader.ReadUInt16(payload, 0);
            ushort substreamType = BiffRecordReader.ReadUInt16(payload, 2);
            if (version != Biff8Version && version != Biff5Version) {
                LegacyXlsUnsupportedFeature feature = BiffUnsupportedRecordDiagnostics.CreateUnsupportedBiffVersionFeature(offset, version, substreamType, sheetName);
                unsupportedFeatures.Add(feature);
                if (BiffUnsupportedRecordDiagnostics.TryCreatePreservedFeatureRecord(feature, payload.Length, out LegacyXlsPreservedFeatureRecord? preservedRecord)) {
                    preservedFeatureRecords.Add(preservedRecord!);
                }

                BiffUnsupportedRecordDiagnostics.AddUnsupportedBiffVersionDiagnostic(diagnostics, offset, version, substreamType, sheetName);
                return false;
            }

            if (expectedSubstreamType.HasValue && substreamType != expectedSubstreamType.Value) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Error,
                    "XLS-BIFF-BOF-SUBSTREAM-UNEXPECTED",
                    $"The BIFF stream starts with substream type 0x{substreamType:X4}, but workbook globals were expected.",
                    sheetName: sheetName,
                    recordOffset: offset,
                    recordType: (ushort)BiffRecordType.Bof,
                    detailCode: $"BiffSubstream:0x{substreamType:X4}"));
                return false;
            }

            return true;
        }
    }
}
