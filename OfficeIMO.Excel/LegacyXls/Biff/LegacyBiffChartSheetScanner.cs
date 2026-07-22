using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Scans legacy chart-sheet substreams into supported chart-sheet and chart metadata.
    /// </summary>
    internal static class LegacyBiffChartSheetScanner {
        internal static void Scan(
            byte[] workbookStream,
            IReadOnlyList<LegacyXlsChartSheet> chartSheets,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsPivotTableRecord> pivotTableRecords,
            List<LegacyXlsChartRecord> chartRecords,
            List<LegacyXlsDrawingRecord> drawingRecords,
            List<LegacyXlsExternalQueryConnection> externalQueryConnections,
            List<LegacyXlsFormulaTokenRecord> formulaTokenRecords,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options,
            LegacyXlsDecodedImageBudget? decodedImageBudget = null) {
            foreach (LegacyXlsChartSheet sheet in chartSheets) {
                if (sheet.StreamOffset <= 0) {
                    continue;
                }

                ScanSheet(workbookStream, sheet, unsupportedFeatures, preservedFeatureRecords, pivotTableRecords, chartRecords, drawingRecords, externalQueryConnections, formulaTokenRecords, externSheets, externalReferences, sheetNames, definedNames, diagnostics, options, decodedImageBudget);
            }
        }

        private static void ScanSheet(
            byte[] workbookStream,
            LegacyXlsChartSheet sheet,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsPivotTableRecord> pivotTableRecords,
            List<LegacyXlsChartRecord> chartRecords,
            List<LegacyXlsDrawingRecord> drawingRecords,
            List<LegacyXlsExternalQueryConnection> externalQueryConnections,
            List<LegacyXlsFormulaTokenRecord> formulaTokenRecords,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options,
            LegacyXlsDecodedImageBudget? decodedImageBudget) {
            if (sheet.StreamOffset >= workbookStream.Length) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-CHART-SHEET-OFFSET-INVALID",
                    $"Chart sheet stream offset {sheet.StreamOffset} is outside the BIFF stream.",
                    sheetName: sheet.Name,
                    recordOffset: sheet.StreamOffset,
                    detailCode: "Sheet:ChartSheet"));
                return;
            }

            int offset = sheet.StreamOffset;
            var chartMetadataState = new BiffChartMetadataReaderState();
            var pivotTableMetadataState = new BiffPivotTableMetadataReaderState();
            while (offset < workbookStream.Length) {
                if (offset + 4 > workbookStream.Length) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-CHART-SHEET-TRUNCATED-HEADER",
                        "A chart-sheet substream ended inside a record header.",
                        sheetName: sheet.Name,
                        recordOffset: offset,
                        detailCode: "Sheet:ChartSheet"));
                    return;
                }

                ushort type = BiffRecordReader.ReadUInt16(workbookStream, offset);
                ushort length = BiffRecordReader.ReadUInt16(workbookStream, offset + 2);
                int payloadOffset = offset + 4;
                if (payloadOffset + length > workbookStream.Length) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-CHART-SHEET-TRUNCATED-PAYLOAD",
                        $"BIFF record 0x{type:X4} declares {length} payload bytes, but the chart-sheet substream ends early.",
                        sheetName: sheet.Name,
                        recordOffset: offset,
                        recordType: type,
                        detailCode: "Sheet:ChartSheet"));
                    return;
                }

                if (type == (ushort)BiffRecordType.Eof) {
                    return;
                }

                if (TryReadChartSheetMetadata(workbookStream, sheet, diagnostics, type, offset, payloadOffset, length, drawingRecords, decodedImageBudget)) {
                    offset = payloadOffset + length;
                    continue;
                }

                if (type != (ushort)BiffRecordType.Bof
                    && TryReadFutureMetadata(workbookStream, sheet, type, offset, payloadOffset, length)) {
                    offset = payloadOffset + length;
                    continue;
                }

                if (type != (ushort)BiffRecordType.Bof
                    && BiffUnsupportedRecordDiagnostics.IsPreserveOnlyFeatureRecord(type)) {
                    byte[] payload = new byte[length];
                    Buffer.BlockCopy(workbookStream, payloadOffset, payload, 0, length);
                    bool isSupportedDrawingMetadata = false;
                    bool isSupportedChartMetadata = false;
                    bool isSupportedPivotTableMetadata = false;

                    var record = new BiffRecord(type, offset, payload);
                    if (BiffDrawingMetadataReader.TryRead(record, sheet.Name, out LegacyXlsDrawingRecord? drawingRecord, decodedImageBudget)) {
                        drawingRecords.Add(drawingRecord!);
                        isSupportedDrawingMetadata = drawingRecord!.HasSupportedDrawingMetadata;
                    }

                    int chartRecordCountBefore = chartRecords.Count;
                    if (BiffChartMetadataReader.TryRead(record, sheet.Name, chartRecords, chartMetadataState, externSheets, externalReferences, sheetNames, definedNames, decodedImageBudget)
                        && chartRecords.Count > chartRecordCountBefore) {
                        BiffChartMetadataReader.ScanFormulaTokens(record, sheet.Name, formulaTokenRecords);
                        sheet.AddChartRecord(chartRecords[chartRecords.Count - 1]);
                        isSupportedChartMetadata = chartRecords[chartRecords.Count - 1].HasSupportedChartMetadata;
                    }

                    int pivotTableRecordCountBefore = pivotTableRecords.Count;
                    if (BiffPivotTableMetadataReader.TryRead(record, sheet.Name, pivotTableRecords, diagnostics, pivotTableMetadataState, formulaTokenRecords)
                        && pivotTableRecords.Count > pivotTableRecordCountBefore) {
                        isSupportedPivotTableMetadata = pivotTableRecords[pivotTableRecords.Count - 1].HasSupportedPivotTableMetadata;
                    }

                    if (BiffExternalQueryConnectionReader.TryRead(record, sheet.Name, diagnostics, out LegacyXlsExternalQueryConnection? externalQueryConnection)) {
                        externalQueryConnections.Add(externalQueryConnection!);
                    }

                    if (!isSupportedDrawingMetadata && !isSupportedChartMetadata && !isSupportedPivotTableMetadata) {
                        LegacyXlsUnsupportedFeature feature = BiffUnsupportedRecordDiagnostics.CreateUnsupportedRecordFeature(type, offset, sheet.Name);
                        unsupportedFeatures.Add(feature);
                        if (BiffUnsupportedRecordDiagnostics.TryCreatePreservedFeatureRecord(feature, length, out LegacyXlsPreservedFeatureRecord? preservedRecord)) {
                            preservedFeatureRecords.Add(preservedRecord!);
                        }

                        if (options.ReportUnsupportedContent) {
                            BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(diagnostics, type, offset, sheet.Name);
                        }
                    }
                }

                offset = payloadOffset + length;
            }
        }

        private static bool TryReadFutureMetadata(
            byte[] workbookStream,
            LegacyXlsChartSheet sheet,
            ushort type,
            int recordOffset,
            int payloadOffset,
            ushort payloadLength) {
            byte[] payload = new byte[payloadLength];
            Buffer.BlockCopy(workbookStream, payloadOffset, payload, 0, payloadLength);
            if (!BiffFutureMetadataReader.TryCreateWorksheetRecord(new BiffRecord(type, recordOffset, payload), out LegacyXlsSheetFutureMetadataRecord? futureMetadataRecord)) {
                return false;
            }

            sheet.AddFutureMetadataRecord(futureMetadataRecord!);
            return true;
        }

        private static bool TryReadChartSheetMetadata(
            byte[] workbookStream,
            LegacyXlsChartSheet sheet,
            List<LegacyXlsImportDiagnostic> diagnostics,
            ushort type,
            int recordOffset,
            int payloadOffset,
            ushort payloadLength,
            List<LegacyXlsDrawingRecord> drawingRecords,
            LegacyXlsDecodedImageBudget? decodedImageBudget) {
            if (type == (ushort)BiffRecordType.Txo) {
                byte[] payload = new byte[payloadLength];
                Buffer.BlockCopy(workbookStream, payloadOffset, payload, 0, payloadLength);
                BiffDrawingMetadataReader.TryRead(new BiffRecord(type, recordOffset, payload), sheet.Name, drawingRecords, decodedImageBudget);
                sheet.IncrementChartTextObjectCount();
                sheet.AddMetadataRecord(LegacyXlsChartSheetMetadataKind.ChartTextObject, recordOffset, type);
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

            sheet.AddMetadataRecord(LegacyXlsChartSheetMetadataKind.ChartPrintSize, recordOffset, type);
            return true;
        }
    }
}
