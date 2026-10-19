using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff;

internal static partial class LegacyBiffWorksheetParser {
    // An embedded chart has its own data cache and metadata. Its NUMBER, BOOLERR,
    // DIMENSIONS and view records must never enter the enclosing worksheet model.
    private static void ReadNestedChartRecord(BiffRecord record, string sheetName, BiffChartMetadataReaderState state,
        IReadOnlyList<BiffExternSheetReference> externSheets, IReadOnlyList<LegacyXlsExternalReference> externalReferences,
        IReadOnlyList<string> sheetNames, IReadOnlyList<string?> names, List<LegacyXlsChartRecord> charts,
        List<LegacyXlsDrawingRecord> drawings, List<LegacyXlsFormulaTokenRecord> tokens,
        List<LegacyXlsUnsupportedFeature> unsupported, List<LegacyXlsPreservedFeatureRecord> preserved,
        List<LegacyXlsImportDiagnostic> diagnostics, LegacyXlsImportOptions options, LegacyXlsDecodedImageBudget? budget) {
        if (record.Type == (ushort)BiffRecordType.Bof) return;
        if (BiffChartMetadataReader.TryRead(record, sheetName, charts, state, externSheets, externalReferences, sheetNames, names, budget)) {
            BiffChartMetadataReader.ScanFormulaTokens(record, sheetName, tokens);
            if (!charts[charts.Count - 1].HasSupportedChartMetadata)
                AddUnsupportedFeature(unsupported, preserved, diagnostics, options, record.Type, record.Offset, sheetName, record.Payload.Length);
        } else if (BiffDrawingMetadataReader.TryRead(record, sheetName, drawings, budget)) {
            if (!drawings[drawings.Count - 1].HasSupportedDrawingMetadata)
                AddUnsupportedFeature(unsupported, preserved, diagnostics, options, record.Type, record.Offset, sheetName, record.Payload.Length);
        } else if (BiffUnsupportedRecordDiagnostics.IsPreserveOnlyFeatureRecord(record.Type)) {
            AddUnsupportedFeature(unsupported, preserved, diagnostics, options, record.Type, record.Offset, sheetName, record.Payload.Length);
        }
    }
}
