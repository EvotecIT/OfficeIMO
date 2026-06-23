using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Scans unsupported sheet substreams for preserve-only BIFF feature records without importing them as worksheets.
    /// </summary>
    internal static class LegacyBiffUnsupportedSheetScanner {
        internal static void Scan(
            byte[] workbookStream,
            IReadOnlyList<LegacyXlsUnsupportedSheet> unsupportedSheets,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsPivotTableRecord> pivotTableRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options) {
            foreach (LegacyXlsUnsupportedSheet sheet in unsupportedSheets) {
                if (!ShouldScan(sheet)) {
                    continue;
                }

                ScanSheet(workbookStream, sheet, unsupportedFeatures, preservedFeatureRecords, pivotTableRecords, diagnostics, options);
            }
        }

        private static bool ShouldScan(LegacyXlsUnsupportedSheet sheet) {
            return sheet.StreamOffset > 0
                && sheet.Kind != LegacyXlsUnsupportedSheetKind.DialogSheet;
        }

        private static void ScanSheet(
            byte[] workbookStream,
            LegacyXlsUnsupportedSheet sheet,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsPivotTableRecord> pivotTableRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options) {
            if (sheet.StreamOffset >= workbookStream.Length) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-UNSUPPORTED-SHEET-OFFSET-INVALID",
                    $"Unsupported sheet stream offset {sheet.StreamOffset} is outside the BIFF stream.",
                    sheetName: sheet.Name,
                    recordOffset: sheet.StreamOffset,
                    detailCode: "Sheet:" + sheet.Kind));
                return;
            }

            int offset = sheet.StreamOffset;
            while (offset < workbookStream.Length) {
                if (offset + 4 > workbookStream.Length) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-UNSUPPORTED-SHEET-TRUNCATED-HEADER",
                        "An unsupported sheet substream ended inside a record header.",
                        sheetName: sheet.Name,
                        recordOffset: offset,
                        detailCode: "Sheet:" + sheet.Kind));
                    return;
                }

                ushort type = BiffRecordReader.ReadUInt16(workbookStream, offset);
                ushort length = BiffRecordReader.ReadUInt16(workbookStream, offset + 2);
                int payloadOffset = offset + 4;
                if (payloadOffset + length > workbookStream.Length) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-UNSUPPORTED-SHEET-TRUNCATED-PAYLOAD",
                        $"BIFF record 0x{type:X4} declares {length} payload bytes, but the unsupported sheet substream ends early.",
                        sheetName: sheet.Name,
                        recordOffset: offset,
                        recordType: type,
                        detailCode: "Sheet:" + sheet.Kind));
                    return;
                }

                if (type == (ushort)BiffRecordType.Eof) {
                    return;
                }

                if (TryReadUnsupportedSheetMetadata(workbookStream, sheet, diagnostics, type, offset, payloadOffset, length)) {
                    offset = payloadOffset + length;
                    continue;
                }

                if (type != (ushort)BiffRecordType.Bof
                    && BiffUnsupportedRecordDiagnostics.IsPreserveOnlyFeatureRecord(type)) {
                    byte[] payload = new byte[length];
                    Buffer.BlockCopy(workbookStream, payloadOffset, payload, 0, length);
                    BiffPivotTableMetadataReader.TryRead(new BiffRecord(type, offset, payload), sheet.Name, pivotTableRecords, diagnostics);
                    LegacyXlsUnsupportedFeature feature = BiffUnsupportedRecordDiagnostics.CreateUnsupportedRecordFeature(type, offset, sheet.Name);
                    unsupportedFeatures.Add(feature);
                    if (BiffUnsupportedRecordDiagnostics.TryCreatePreservedFeatureRecord(feature, length, out LegacyXlsPreservedFeatureRecord? preservedRecord)) {
                        preservedFeatureRecords.Add(preservedRecord!);
                    }

                    if (options.ReportUnsupportedRecords) {
                        BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(diagnostics, type, offset, sheet.Name);
                    }
                }

                offset = payloadOffset + length;
            }
        }

        private static bool TryReadUnsupportedSheetMetadata(
            byte[] workbookStream,
            LegacyXlsUnsupportedSheet sheet,
            List<LegacyXlsImportDiagnostic> diagnostics,
            ushort type,
            int recordOffset,
            int payloadOffset,
            ushort payloadLength) {
            if (sheet.Kind != LegacyXlsUnsupportedSheetKind.ChartSheet) {
                return false;
            }

            if (type == (ushort)BiffRecordType.Txo) {
                sheet.IncrementChartTextObjectCount();
                sheet.AddMetadataRecord(LegacyXlsUnsupportedSheetMetadataKind.ChartTextObject, recordOffset, type);
                return true;
            }

            if (type != (ushort)BiffRecordType.PrintSize) {
                return false;
            }

            if (payloadLength < 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-CHART-METADATA-SHORT",
                    "The chart sheet PrintSize record is shorter than expected.",
                    sheetName: sheet.Name,
                    recordOffset: recordOffset,
                    recordType: type));
            } else {
                ushort printSize = BiffRecordReader.ReadUInt16(workbookStream, payloadOffset);
                if (printSize > 3) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-CHART-PRINTSIZE-UNEXPECTED",
                        $"The chart sheet PrintSize record contains unexpected value {printSize}.",
                        sheetName: sheet.Name,
                        recordOffset: recordOffset,
                        recordType: type));
                }

                sheet.SetChartPrintSize(printSize);
            }

            sheet.AddMetadataRecord(LegacyXlsUnsupportedSheetMetadataKind.ChartPrintSize, recordOffset, type);
            return true;
        }
    }
}
