using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class LegacyBiffWorksheetParser {
        internal static void Parse(
            byte[] workbookStream,
            ushort workbookGlobalsBiffVersion,
            LegacyXlsWorkbook workbook,
            LegacyXlsWorksheet sheet,
            IReadOnlyList<BiffStringReader.BiffStringValue> sharedStrings,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsPivotTableRecord> pivotTableRecords,
            List<LegacyXlsChartRecord> chartRecords,
            List<LegacyXlsDrawingRecord> drawingRecords,
            List<LegacyXlsExternalQueryConnection> externalQueryConnections,
            IReadOnlyList<LegacyXlsDifferentialFormat> differentialFormats,
            LegacyXlsCalculationSettings calculationSettings,
            List<LegacyXlsFormulaTokenRecord> formulaTokenRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options,
            LegacyXlsDecodedImageBudget? decodedImageBudget = null) {
            if (sheet.StreamOffset < 0 || sheet.StreamOffset >= workbookStream.Length) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-SHEET-OFFSET-INVALID",
                    $"Worksheet stream offset {sheet.StreamOffset} is outside the BIFF stream.",
                    sheetName: sheet.Name,
                    recordOffset: sheet.StreamOffset));
                return;
            }

            int offset = sheet.StreamOffset;
            bool frozenWindow = false;
            int nestedSubstreamDepth = 0;
            PendingFormulaString? pendingFormulaString = null;
            var commentState = new BiffCommentImportState(sheet, decodedImageBudget);
            var drawingContinuationState = new BiffDrawingContinuationImportState(sheet.Name, drawingRecords, decodedImageBudget);
            var drawingTextObjectState = new BiffDrawingTextObjectImportState(sheet.Name, drawingRecords, decodedImageBudget);
            var conditionalFormattingState = new BiffConditionalFormattingImportState(workbook, sheet, externSheets, externalReferences, sheetNames, definedNames, differentialFormats);
            var sharedFormulaState = new BiffSharedFormulaImportState(sheet, externSheets, externalReferences, sheetNames, definedNames, formulaTokenRecords, diagnostics, options);
            var chartMetadataState = new BiffChartMetadataReaderState();
            var pivotTableMetadataState = new BiffPivotTableMetadataReaderState();
            while (offset < workbookStream.Length) {
                if (offset + 4 > workbookStream.Length) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Error,
                        "XLS-BIFF-SHEET-TRUNCATED-HEADER",
                        "The worksheet stream ended inside a record header.",
                        sheetName: sheet.Name,
                        recordOffset: offset));
                    return;
                }

                ushort type = BiffRecordReader.ReadUInt16(workbookStream, offset);
                ushort length = BiffRecordReader.ReadUInt16(workbookStream, offset + 2);
                int payloadOffset = offset + 4;
                if (payloadOffset + length > workbookStream.Length) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Error,
                        "XLS-BIFF-SHEET-TRUNCATED-PAYLOAD",
                        $"BIFF record 0x{type:X4} declares {length} payload bytes, but the worksheet stream ends early.",
                        sheetName: sheet.Name,
                        recordOffset: offset,
                        recordType: type));
                    return;
                }

                byte[] payload = new byte[length];
                Buffer.BlockCopy(workbookStream, payloadOffset, payload, 0, length);
                if (offset == sheet.StreamOffset) {
                    if (type != (ushort)BiffRecordType.Bof) {
                        diagnostics.Add(new LegacyXlsImportDiagnostic(
                            LegacyXlsDiagnosticSeverity.Error,
                            "XLS-BIFF-SHEET-BOF-MISSING",
                            "The worksheet substream does not start with a BOF record.",
                            sheetName: sheet.Name,
                            recordOffset: offset,
                            recordType: type));
                        return;
                    }

                    if (!LegacyBiffVersionValidator.ValidateWorksheetBof(payload, offset, sheet.Name, unsupportedFeatures, preservedFeatureRecords, diagnostics)) {
                        return;
                    }
                } else if (type == (ushort)BiffRecordType.Bof) {
                    nestedSubstreamDepth++;
                }

                if (type == (ushort)BiffRecordType.Eof) {
                    if (nestedSubstreamDepth > 0) {
                        nestedSubstreamDepth--;
                        offset = payloadOffset + length;
                        continue;
                    }

                    FlushPendingFormulaString(sheet, sharedFormulaState, ref pendingFormulaString);
                    drawingContinuationState.AddUnresolvedFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
                    commentState.AddPendingDrawingFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
                    commentState.AddUnresolvedFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
                    drawingTextObjectState.AddUnresolvedFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
                    conditionalFormattingState.AddUnresolvedFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
                    sharedFormulaState.AddUnresolvedDiagnostics();
                    return;
                }

                ParseWorksheetRecord(sheet, workbookGlobalsBiffVersion, sharedStrings, externSheets, externalReferences, sheetNames, definedNames, unsupportedFeatures, preservedFeatureRecords, pivotTableRecords, chartRecords, drawingRecords, externalQueryConnections, calculationSettings, formulaTokenRecords, diagnostics, options, decodedImageBudget, commentState, drawingTextObjectState, drawingContinuationState, conditionalFormattingState, sharedFormulaState, chartMetadataState, pivotTableMetadataState, type, offset, payload, ref frozenWindow, ref pendingFormulaString);
                offset = payloadOffset + length;
            }

            FlushPendingFormulaString(sheet, sharedFormulaState, ref pendingFormulaString);
            drawingContinuationState.AddUnresolvedFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
            commentState.AddPendingDrawingFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
            commentState.AddUnresolvedFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
            drawingTextObjectState.AddUnresolvedFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
            conditionalFormattingState.AddUnresolvedFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
            sharedFormulaState.AddUnresolvedDiagnostics();
        }

        private static void ParseWorksheetRecord(
            LegacyXlsWorksheet sheet,
            ushort workbookGlobalsBiffVersion,
            IReadOnlyList<BiffStringReader.BiffStringValue> sharedStrings,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsPivotTableRecord> pivotTableRecords,
            List<LegacyXlsChartRecord> chartRecords,
            List<LegacyXlsDrawingRecord> drawingRecords,
            List<LegacyXlsExternalQueryConnection> externalQueryConnections,
            LegacyXlsCalculationSettings calculationSettings,
            List<LegacyXlsFormulaTokenRecord> formulaTokenRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options,
            LegacyXlsDecodedImageBudget? decodedImageBudget,
            BiffCommentImportState commentState,
            BiffDrawingTextObjectImportState drawingTextObjectState,
            BiffDrawingContinuationImportState drawingContinuationState,
            BiffConditionalFormattingImportState conditionalFormattingState,
            BiffSharedFormulaImportState sharedFormulaState,
            BiffChartMetadataReaderState chartMetadataState,
            BiffPivotTableMetadataReaderState pivotTableMetadataState,
            ushort type,
            int offset,
            byte[] payload,
            ref bool frozenWindow,
            ref PendingFormulaString? pendingFormulaString) {
            try {
                if (pendingFormulaString != null
                    && type != (ushort)BiffRecordType.String
                    && (type != (ushort)BiffRecordType.Continue || !pendingFormulaString.HasStringPayload)) {
                    FlushPendingFormulaString(sheet, sharedFormulaState, ref pendingFormulaString);
                }

                if (type != (ushort)BiffRecordType.Continue) {
                    drawingContinuationState.AddUnresolvedFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
                }

                if (type != (ushort)BiffRecordType.Drawing && type != (ushort)BiffRecordType.Obj && type != (ushort)BiffRecordType.Continue) {
                    commentState.AddPendingDrawingFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
                }

                switch ((BiffRecordType)type) {
                    case BiffRecordType.Blank:
                        if (payload.Length >= 6) {
                            sheet.AddCell(new LegacyXlsCell(
                                BiffRecordReader.ReadUInt16(payload, 0) + 1,
                                BiffRecordReader.ReadUInt16(payload, 2) + 1,
                                LegacyXlsCellValueKind.Blank,
                                null,
                                BiffRecordReader.ReadUInt16(payload, 4)));
                        }

                        break;
                    case BiffRecordType.BottomMargin:
                        ParseMargin(sheet, payload, BiffRecordType.BottomMargin);
                        break;
                    case BiffRecordType.ColInfo:
                        ParseColInfo(sheet, payload);
                        break;
                    case BiffRecordType.Country:
                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.Country, offset, type);
                        break;
                    case BiffRecordType.CodeName:
                        if (BiffCodeNameReader.TryRead(new BiffRecord(type, offset, payload), sheet.Name, diagnostics, out string? sheetCodeName)) {
                            sheet.SetCodeName(sheetCodeName);
                        }

                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.CodeName, offset, type);
                        break;
                    case BiffRecordType.CondFmt:
                        if (!conditionalFormattingState.TryReadHeader(payload, offset)) {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.Cf:
                        if (!conditionalFormattingState.TryReadRule(payload, offset, out BiffFormulaReadFailure? conditionalFormattingFormulaFailure)) {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                            AddFormulaTokenDiagnostic(diagnostics, options, conditionalFormattingFormulaFailure, type, offset, sheet.Name, "ConditionalFormatting", "Conditional-formatting formula");
                        }

                        break;
                    case BiffRecordType.CfEx:
                        if (!conditionalFormattingState.TryReadExtension(payload, offset, out bool hasUnprojectedFormatting) || hasUnprojectedFormatting) {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.CellWatch:
                        if (BiffCellWatchReader.TryRead(new BiffRecord(type, offset, payload), out LegacyXlsCellWatch? cellWatch)) {
                            sheet.AddCellWatch(cellWatch!);
                            sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.CellWatches, offset, type);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.Continue:
                        if (pendingFormulaString?.HasStringPayload == true) {
                            pendingFormulaString.AddStringPayload(payload);
                            break;
                        }

                        if (!commentState.TryReadContinue(payload)
                            && !drawingTextObjectState.TryReadContinue(payload)) {
                            if (drawingContinuationState.TryReadContinue(
                                payload,
                                unsupportedFeatures,
                                preservedFeatureRecords,
                                diagnostics,
                                options.ReportUnsupportedContent,
                                out BiffRecord? assembledDrawingRecord)
                                && assembledDrawingRecord.HasValue) {
                                commentState.TryReadDrawingAnchors(assembledDrawingRecord.Value, out _);
                            }
                        }

                        break;
                    case BiffRecordType.DCon:
                        if (BiffDataConsolidationSettingsReader.TryRead(new BiffRecord(type, offset, payload), out LegacyXlsDataConsolidationSettings? dataConsolidationSettings)) {
                            sheet.SetDataConsolidationSettings(dataConsolidationSettings!);
                            sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.DataConsolidation, offset, type);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.DefColWidth:
                        ParseDefaultColumnWidth(sheet, payload);
                        break;
                    case BiffRecordType.AutoFilterInfo:
                        if (BiffAutoFilterReader.TryReadInfo(payload, out ushort dropDownCount)) {
                            sheet.SetAutoFilterDropDownCount(dropDownCount);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.AutoFilter:
                        if (BiffAutoFilterReader.TryReadCriteria(payload, out LegacyXlsAutoFilterCriteria? criteria)) {
                            sheet.AddAutoFilterCriteria(criteria!);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.Array:
                        if (!sharedFormulaState.TryConsumeArrayFormula(payload, offset)) {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.FilterMode:
                        break;
                    case BiffRecordType.FeatHdr:
                        if (BiffWorksheetProtectionFeatureReader.TryRead(new BiffRecord(type, offset, payload), out LegacyXlsWorksheetProtectionPermissions? permissions)) {
                            sheet.SetProtectionPermissions(permissions!);
                        } else if (BiffIgnoredErrorsFeatureReader.TryReadHeader(new BiffRecord(type, offset, payload))) {
                            sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.IgnoredErrors, offset, type);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.Feat:
                        if (BiffProtectedRangeFeatureReader.TryRead(new BiffRecord(type, offset, payload), out LegacyXlsProtectedRange? protectedRange)) {
                            sheet.AddProtectedRange(protectedRange!);
                        } else if (BiffIgnoredErrorsFeatureReader.TryRead(new BiffRecord(type, offset, payload), out LegacyXlsIgnoredError? ignoredError)) {
                            sheet.AddIgnoredError(ignoredError!);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.DefaultRowHeight:
                        ParseDefaultRowHeight(sheet, payload);
                        break;
                    case BiffRecordType.DbCell:
                        if (workbookGlobalsBiffVersion == LegacyBiffVersionValidator.Biff5Version) {
                            sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.RowBlockIndex, offset, type);
                        } else if (BiffPivotTableMetadataReader.TryRead(new BiffRecord(type, offset, payload), sheet.Name, pivotTableRecords, diagnostics, pivotTableMetadataState, formulaTokenRecords)) {
                            LegacyXlsPivotTableRecord pivotTableRecord = pivotTableRecords[pivotTableRecords.Count - 1];
                            if (!pivotTableRecord.HasSupportedPivotTableMetadata) {
                                AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                            }
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.Dimensions:
                        ParseDimensions(sheet, payload);
                        break;
                    case BiffRecordType.DVal:
                        if (BiffDataValidationReader.TryReadCollectionHeader(new BiffRecord(type, offset, payload), sheet.Name, out LegacyXlsDataValidationCollectionRecord? collectionRecord)) {
                            sheet.AddDataValidationCollection(collectionRecord!);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.Dv:
                        if (BiffDataValidationReader.TryRead(payload, externSheets, externalReferences, sheetNames, definedNames, out LegacyXlsDataValidation? validation, out BiffFormulaReadFailure? dataValidationFormulaFailure)) {
                            sheet.AddDataValidation(validation!);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                            AddFormulaTokenDiagnostic(diagnostics, options, dataValidationFormulaFailure, type, offset, sheet.Name, "DataValidation", "Data-validation formula");
                        }

                        break;
                    case BiffRecordType.BoolErr:
                        if (payload.Length >= 8) {
                            bool isError = payload[7] != 0;
                            sheet.AddCell(new LegacyXlsCell(
                                BiffRecordReader.ReadUInt16(payload, 0) + 1,
                                BiffRecordReader.ReadUInt16(payload, 2) + 1,
                                isError ? LegacyXlsCellValueKind.Error : LegacyXlsCellValueKind.Boolean,
                                isError ? BiffErrorValue.ToText(payload[6]) : payload[6] != 0,
                                BiffRecordReader.ReadUInt16(payload, 4)));
                        }

                        break;
                    case BiffRecordType.CalcCount:
                    case BiffRecordType.CalcMode:
                    case BiffRecordType.CalcPrecision:
                    case BiffRecordType.CalcRefMode:
                    case BiffRecordType.CalcDelta:
                    case BiffRecordType.CalcIter:
                    case BiffRecordType.CalcSaveRecalc:
                        BiffCalculationSettingsReader.TryRead(new BiffRecord(type, offset, payload), sheet.Name, calculationSettings, diagnostics);
                        break;
                    case BiffRecordType.Uncalced:
                        ParseUncalced(sheet, payload, offset, type, diagnostics);
                        break;
                    case BiffRecordType.PhoneticInfo:
                        if (!TryParsePhoneticInfo(sheet, payload, offset, type, diagnostics, out _)) {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.Formula:
                        ParseFormula(sheet, payload, externSheets, externalReferences, sheetNames, definedNames, formulaTokenRecords, diagnostics, options, sharedFormulaState, offset, ref pendingFormulaString);
                        break;
                    case BiffRecordType.Footer:
                        ParseHeaderFooter(sheet, payload, isHeader: false);
                        break;
                    case BiffRecordType.GridSet:
                        sheet.SetGridSet(BiffWorksheetMetadataReader.ReadGridSet(new BiffRecord(type, offset, payload), sheet.Name, diagnostics));
                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.GridSet, offset, type);
                        break;
                    case BiffRecordType.Gcw:
                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.ColumnDisplay, offset, type);
                        break;
                    case BiffRecordType.Guts:
                        (byte rowLevel, byte columnLevel) = BiffWorksheetMetadataReader.ReadGuts(payload);
                        sheet.SetOutlineLevels(rowLevel, columnLevel);
                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.OutlineLevels, offset, type);
                        break;
                    case BiffRecordType.Header:
                        ParseHeaderFooter(sheet, payload, isHeader: true);
                        break;
                    case BiffRecordType.HeaderFooter:
                        if (!TryParseHeaderFooterExtension(sheet, payload, offset)) {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.HCenter:
                    case BiffRecordType.PrintGrid:
                    case BiffRecordType.PrintRowCol:
                    case BiffRecordType.VCenter:
                        ParsePrintOption(sheet, payload, (BiffRecordType)type);
                        break;
                    case BiffRecordType.HorizontalPageBreaks:
                        ParseHorizontalPageBreaks(sheet, payload);
                        break;
                    case BiffRecordType.HLink:
                        if (BiffHyperlinkReader.TryRead(payload, out LegacyXlsHyperlink? hyperlink)) {
                            sheet.AddHyperlink(hyperlink!);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.HLinkTooltip:
                        if (BiffHyperlinkTooltipReader.TryRead(payload, out BiffHyperlinkTooltip? tooltip)
                            && sheet.TrySetHyperlinkTooltip(tooltip!.StartRow, tooltip.StartColumn, tooltip.EndRow, tooltip.EndColumn, tooltip.Text)) {
                            sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.Hyperlink, offset, type);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.Index:
                        sheet.SetRowBlockIndex(BiffWorksheetMetadataReader.ReadIndex(payload, workbookGlobalsBiffVersion == LegacyBiffVersionValidator.Biff5Version));
                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.RowBlockIndex, offset, type);
                        break;
                    case BiffRecordType.Label:
                        if (payload.Length >= 8) {
                            int stringOffset = 6;
                            BiffStringReader.BiffStringValue value = ReadLabelValue(payload, ref stringOffset);
                            sheet.AddCell(new LegacyXlsCell(
                                BiffRecordReader.ReadUInt16(payload, 0) + 1,
                                BiffRecordReader.ReadUInt16(payload, 2) + 1,
                                LegacyXlsCellValueKind.Text,
                                value.Text,
                                BiffRecordReader.ReadUInt16(payload, 4),
                                textFormattingRuns: value.FormattingRuns));
                        }

                        break;
                    case BiffRecordType.LabelSst:
                        if (payload.Length >= 10) {
                            uint sharedStringIndex = BiffRecordReader.ReadUInt32(payload, 6);
                            BiffStringReader.BiffStringValue value = sharedStringIndex < sharedStrings.Count
                                ? sharedStrings[(int)sharedStringIndex]
                                : new BiffStringReader.BiffStringValue(string.Empty);
                            sheet.AddCell(new LegacyXlsCell(
                                BiffRecordReader.ReadUInt16(payload, 0) + 1,
                                BiffRecordReader.ReadUInt16(payload, 2) + 1,
                                LegacyXlsCellValueKind.Text,
                                value.Text,
                                BiffRecordReader.ReadUInt16(payload, 4),
                                textFormattingRuns: value.FormattingRuns));
                        }

                        break;
                    case BiffRecordType.LeftMargin:
                        ParseMargin(sheet, payload, BiffRecordType.LeftMargin);
                        break;
                    case BiffRecordType.MergeCells:
                        ParseMergeCells(sheet, payload);
                        break;
                    case BiffRecordType.MulBlank:
                        ParseMulBlank(sheet, payload);
                        break;
                    case BiffRecordType.MulRk:
                        ParseMulRk(sheet, payload);
                        break;
                    case BiffRecordType.Number:
                        if (payload.Length >= 14) {
                            sheet.AddCell(new LegacyXlsCell(
                                BiffRecordReader.ReadUInt16(payload, 0) + 1,
                                BiffRecordReader.ReadUInt16(payload, 2) + 1,
                                LegacyXlsCellValueKind.Number,
                                BiffRecordReader.ReadDouble(payload, 6),
                                BiffRecordReader.ReadUInt16(payload, 4)));
                        }

                        break;
                    case BiffRecordType.Note:
                        if (!commentState.TryReadNote(payload, offset)) {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.ObjProtect:
                        ParseObjectProtection(sheet, payload);
                        break;
                    case BiffRecordType.Obj:
                        if (!commentState.TryReadObject(payload)) {
                            commentState.AddPendingDrawingFeatures(unsupportedFeatures, preservedFeatureRecords, diagnostics, options.ReportUnsupportedContent);
                            if (BiffDrawingMetadataReader.TryRead(new BiffRecord(type, offset, payload), sheet.Name, out LegacyXlsDrawingRecord? drawingRecord, decodedImageBudget)) {
                                drawingRecords.Add(drawingRecord!);
                                if (!drawingRecord!.HasSupportedDrawingMetadata) {
                                    AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                                }
                            } else {
                                AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                            }
                        }

                        break;
                    case BiffRecordType.Pane:
                        if (frozenWindow) {
                            ParseFrozenPane(sheet, payload);
                        } else if (TryParseSplitPane(sheet, payload)) {
                            sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.Pane, offset, type);
                        } else if (HasSplitPane(payload)) {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.Plv:
                        var pageLayoutViewRecord = new BiffRecord(type, offset, payload);
                        if (BiffFutureMetadataReader.TryCreateWorksheetRecord(pageLayoutViewRecord, out LegacyXlsSheetFutureMetadataRecord? pageLayoutViewMetadataRecord)) {
                            sheet.AddFutureMetadataRecord(pageLayoutViewMetadataRecord!);
                        }

                        if (BiffPageLayoutViewReader.TryRead(pageLayoutViewRecord, out uint? pageLayoutZoomScale)) {
                            sheet.SetPageLayoutView(true, pageLayoutZoomScale);
                        }

                        break;
                    case BiffRecordType.Password:
                        ParsePassword(sheet, payload);
                        break;
                    case BiffRecordType.Pls:
                        BiffPrinterSettingsReader.Validate(new BiffRecord(type, offset, payload), sheet.Name, diagnostics);
                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.PrinterSettings, offset, type);
                        break;
                    case BiffRecordType.PrintSize:
                        ParsePrintSize(sheet, payload, diagnostics, offset);
                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.PrintSize, offset, type);
                        break;
                    case BiffRecordType.Protect:
                        ParseProtect(sheet, payload);
                        break;
                    case BiffRecordType.RightMargin:
                        ParseMargin(sheet, payload, BiffRecordType.RightMargin);
                        break;
                    case BiffRecordType.Row:
                        ParseRow(sheet, payload);
                        break;
                    case BiffRecordType.Scl:
                        ParseZoomScale(sheet, payload);
                        break;
                    case BiffRecordType.ScenMan:
                        if (BiffScenarioReader.TryReadManager(new BiffRecord(type, offset, payload), out LegacyXlsScenarioManager? scenarioManager)) {
                            sheet.SetScenarioManager(scenarioManager!);
                            sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.Scenarios, offset, type);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.Scenario:
                        if (BiffScenarioReader.TryReadScenario(new BiffRecord(type, offset, payload), out LegacyXlsScenario? scenario)) {
                            sheet.AddScenario(scenario!);
                            sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.Scenarios, offset, type);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.ScenarioProtect:
                        ParseScenarioProtection(sheet, payload);
                        break;
                    case BiffRecordType.SheetExt:
                        var sheetExtensionRecord = new BiffRecord(type, offset, payload);
                        if (BiffFutureMetadataReader.TryCreateWorksheetRecord(sheetExtensionRecord, out LegacyXlsSheetFutureMetadataRecord? sheetExtensionMetadataRecord)) {
                            sheet.AddFutureMetadataRecord(sheetExtensionMetadataRecord!);
                        }

                        if (!TryParseSheetExtension(sheet, payload, out bool hasUnsupportedSheetExtensionMetadata)
                            || hasUnsupportedSheetExtensionMetadata) {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.Selection:
                        sheet.AddSelection(BiffWorksheetMetadataReader.ReadSelection(payload));
                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.Selection, offset, type);
                        break;
                    case BiffRecordType.Setup:
                        ParseSetup(sheet, payload);
                        break;
                    case BiffRecordType.Sort:
                        ParseSortSettings(sheet, payload, unsupportedFeatures, diagnostics, options, offset);
                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.Sort, offset, type);
                        break;
                    case BiffRecordType.Rk:
                        if (payload.Length >= 10) {
                            sheet.AddCell(new LegacyXlsCell(
                                BiffRecordReader.ReadUInt16(payload, 0) + 1,
                                BiffRecordReader.ReadUInt16(payload, 2) + 1,
                                LegacyXlsCellValueKind.Number,
                                BiffRkNumberReader.ReadRkNumber(BiffRecordReader.ReadUInt32(payload, 6)),
                                BiffRecordReader.ReadUInt16(payload, 4)));
                        }

                        break;
                    case BiffRecordType.String:
                        if (pendingFormulaString != null) {
                            pendingFormulaString.AddStringPayload(payload);
                        } else if (options.ReportUnsupportedContent) {
                            diagnostics.Add(new LegacyXlsImportDiagnostic(
                                LegacyXlsDiagnosticSeverity.Info,
                                "XLS-BIFF-FORMULA-STRING-ORPHANED",
                                "A formula String record was found without a pending Formula record.",
                                sheetName: sheet.Name,
                                recordOffset: offset,
                                recordType: type));
                        }

                        break;
                    case BiffRecordType.ShrFmla:
                        if (!sharedFormulaState.TryReadDefinition(payload, offset)) {
                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                    case BiffRecordType.TopMargin:
                        ParseMargin(sheet, payload, BiffRecordType.TopMargin);
                        break;
                    case BiffRecordType.Txo:
                        if (!commentState.TryReadTextObject(payload)) {
                            if (!drawingTextObjectState.TryReadTextObject(new BiffRecord(type, offset, payload))) {
                                BiffDrawingMetadataReader.TryRead(new BiffRecord(type, offset, payload), sheet.Name, drawingRecords, decodedImageBudget);
                                AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                            }
                        }

                        break;
                    case BiffRecordType.VerticalPageBreaks:
                        ParseVerticalPageBreaks(sheet, payload);
                        break;
                    case BiffRecordType.Window2:
                        ParseWindow2(sheet, payload, out frozenWindow);
                        break;
                    case BiffRecordType.Window1:
                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.Window, offset, type);
                        break;
                    case BiffRecordType.WsBool:
                        sheet.SetSheetOptions(BiffWorksheetMetadataReader.ReadWsBool(payload));
                        sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.SheetOptions, offset, type);
                        break;
                    default:
                        if (type != (ushort)BiffRecordType.Bof) {
                            if (BiffFutureMetadataReader.TryCreateWorksheetRecord(new BiffRecord(type, offset, payload), out LegacyXlsSheetFutureMetadataRecord? futureMetadataRecord)) {
                                sheet.AddFutureMetadataRecord(futureMetadataRecord!);
                                break;
                            }

                            var record = new BiffRecord(type, offset, payload);
                            if (BiffDrawingContinuationImportState.RequiresContinuation(record)
                                && drawingContinuationState.TryReadDrawing(
                                    record,
                                    unsupportedFeatures,
                                    preservedFeatureRecords,
                                    diagnostics,
                                    options.ReportUnsupportedContent)) {
                                if (type == (ushort)BiffRecordType.Drawing) {
                                    commentState.TryReadDrawingAnchors(record, out _);
                                }

                                break;
                            }

                            if (type == (ushort)BiffRecordType.Drawing
                                && commentState.TryReadDrawingAnchors(record, out LegacyXlsDrawingRecord? commentDrawingRecord)) {
                                drawingRecords.Add(commentDrawingRecord!);
                                break;
                            }

                            if (drawingContinuationState.TryReadDrawing(
                                record,
                                unsupportedFeatures,
                                preservedFeatureRecords,
                                diagnostics,
                                options.ReportUnsupportedContent)) {
                                break;
                            }

                            if (BiffTableDefinitionReader.IsTableDefinitionRecord(type)
                                && BiffTableDefinitionReader.TryRead(record, sheet, diagnostics, out bool tableDefinitionProjectable)) {
                                if (!tableDefinitionProjectable) {
                                    AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                                }

                                break;
                            }

                            if (BiffDrawingMetadataReader.TryRead(new BiffRecord(type, offset, payload), sheet.Name, drawingRecords, decodedImageBudget)) {
                                LegacyXlsDrawingRecord drawingRecord = drawingRecords[drawingRecords.Count - 1];
                                if (!drawingRecord.HasSupportedDrawingMetadata) {
                                    AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                                }

                                break;
                            }

                            var chartRecord = new BiffRecord(type, offset, payload);
                            if (BiffChartMetadataReader.TryRead(chartRecord, sheet.Name, chartRecords, chartMetadataState, externSheets, externalReferences, sheetNames, definedNames, decodedImageBudget)) {
                                BiffChartMetadataReader.ScanFormulaTokens(chartRecord, sheet.Name, formulaTokenRecords);
                                LegacyXlsChartRecord parsedChartRecord = chartRecords[chartRecords.Count - 1];
                                if (!parsedChartRecord.HasSupportedChartMetadata) {
                                    AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                                }

                                break;
                            }

                            if (BiffPivotTableMetadataReader.TryRead(new BiffRecord(type, offset, payload), sheet.Name, pivotTableRecords, diagnostics, pivotTableMetadataState, formulaTokenRecords)) {
                                LegacyXlsPivotTableRecord pivotTableRecord = pivotTableRecords[pivotTableRecords.Count - 1];
                                if (!pivotTableRecord.HasSupportedPivotTableMetadata) {
                                    AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                                }

                                break;
                            }

                            if (BiffExternalQueryConnectionReader.TryRead(new BiffRecord(type, offset, payload), sheet.Name, diagnostics, out LegacyXlsExternalQueryConnection? externalQueryConnection)) {
                                externalQueryConnections.Add(externalQueryConnection!);
                                break;
                            }

                            AddUnsupportedFeature(unsupportedFeatures, preservedFeatureRecords, diagnostics, options, type, offset, sheet.Name, payload.Length);
                        }

                        break;
                }
            } catch (Exception ex) when (ex is InvalidDataException || ex is OverflowException) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-SHEET-RECORD-INVALID",
                    $"BIFF record 0x{type:X4} could not be parsed. {ex.Message}",
                    sheetName: sheet.Name,
                    recordOffset: offset,
                    recordType: type));
            }
        }

        private static void AddUnsupportedFeature(
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options,
            ushort type,
            int offset,
            string? sheetName,
            int payloadLength) {
            LegacyXlsUnsupportedFeature feature = BiffUnsupportedRecordDiagnostics.CreateUnsupportedRecordFeature(type, offset, sheetName);
            unsupportedFeatures.Add(feature);
            if (BiffUnsupportedRecordDiagnostics.TryCreatePreservedFeatureRecord(feature, payloadLength, out LegacyXlsPreservedFeatureRecord? preservedRecord)) {
                preservedFeatureRecords.Add(preservedRecord!);
            }

            if (options.ReportUnsupportedContent) {
                BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(diagnostics, type, offset, sheetName);
            }
        }

        private static void AddFormulaTokenDiagnostic(
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options,
            BiffFormulaReadFailure? failure,
            ushort type,
            int offset,
            string? sheetName,
            string formulaContext,
            string formulaContextDisplay) {
            if (!options.ReportUnsupportedContent || failure == null) {
                return;
            }

            diagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Info,
                "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED",
                $"{failure.Description} {formulaContextDisplay} was preserved without projection.",
                sheetName: sheetName,
                recordOffset: offset,
                recordType: type,
                detailCode: failure.DetailCode,
                formulaContext: formulaContext,
                formulaToken: failure.Token,
                formulaTokenName: failure.TokenName,
                formulaTokenOffset: failure.TokenOffset));
        }

        private static void ParseZoomScale(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 4) {
                return;
            }

            short numerator = BiffRecordReader.ReadInt16(payload, 0);
            short denominator = BiffRecordReader.ReadInt16(payload, 2);
            if (numerator <= 0 || denominator <= 0) {
                return;
            }

            double scale = numerator * 100d / denominator;
            if (scale < 10d || scale > 400d) {
                return;
            }

            sheet.SetZoomScale((uint)Math.Round(scale, MidpointRounding.AwayFromZero));
        }

        private static void ParsePrintSize(
            LegacyXlsWorksheet sheet,
            byte[] payload,
            List<LegacyXlsImportDiagnostic> diagnostics,
            int offset) {
            if (payload.Length < 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-WORKSHEET-PRINTSIZE-SHORT",
                    "The worksheet PrintSize record is shorter than expected.",
                    sheetName: sheet.Name,
                    recordOffset: offset,
                    recordType: (ushort)BiffRecordType.PrintSize));
                return;
            }

            ushort printSize = BiffRecordReader.ReadUInt16(payload, 0);
            if (printSize > 3) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-WORKSHEET-PRINTSIZE-UNEXPECTED",
                    $"The worksheet PrintSize record contains unexpected value {printSize}.",
                    sheetName: sheet.Name,
                    recordOffset: offset,
                    recordType: (ushort)BiffRecordType.PrintSize));
            }

            sheet.GetOrCreatePageSetup().PrintedSize = printSize;
        }

        private static void ParseDimensions(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length == 10) {
                ushort biff5FirstRow = BiffRecordReader.ReadUInt16(payload, 0);
                ushort biff5RowAfterLast = BiffRecordReader.ReadUInt16(payload, 2);
                ushort biff5FirstColumn = BiffRecordReader.ReadUInt16(payload, 4);
                ushort biff5ColumnAfterLast = BiffRecordReader.ReadUInt16(payload, 6);
                SetDimensions(sheet, biff5FirstRow, biff5RowAfterLast, biff5FirstColumn, biff5ColumnAfterLast);
                return;
            }

            if (payload.Length < 14) {
                return;
            }

            uint firstRow = BiffRecordReader.ReadUInt32(payload, 0);
            uint rowAfterLast = BiffRecordReader.ReadUInt32(payload, 4);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, 8);
            ushort columnAfterLast = BiffRecordReader.ReadUInt16(payload, 10);
            SetDimensions(sheet, firstRow, rowAfterLast, firstColumn, columnAfterLast);
        }

        private static void SetDimensions(LegacyXlsWorksheet sheet, uint firstRow, uint rowAfterLast, ushort firstColumn, ushort columnAfterLast) {
            if (rowAfterLast == 0 || columnAfterLast == 0) {
                sheet.SetDeclaredUsedRange(LegacyXlsWorksheetDimension.Empty);
                return;
            }

            if (rowAfterLast > 0x00010000U || columnAfterLast > 0x0100U || rowAfterLast <= firstRow || columnAfterLast <= firstColumn) {
                throw new InvalidDataException("The DIMENSIONS record contains invalid used-range bounds.");
            }

            sheet.SetDeclaredUsedRange(LegacyXlsWorksheetDimension.FromOneBasedBounds(
                checked((int)firstRow + 1),
                firstColumn + 1,
                checked((int)rowAfterLast),
                columnAfterLast));
        }

        private static BiffStringReader.BiffStringValue ReadLabelValue(byte[] payload, ref int stringOffset) {
            int originalOffset = stringOffset;
            try {
                return BiffStringReader.ReadUnicodeStringValue(payload, ref stringOffset);
            } catch (InvalidDataException) {
                stringOffset = originalOffset;
                return new BiffStringReader.BiffStringValue(BiffStringReader.ReadByteString(payload, ref stringOffset));
            }
        }

        private static void ParseHorizontalPageBreaks(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 2) {
                return;
            }

            ushort count = BiffRecordReader.ReadUInt16(payload, 0);
            if (count > 1026) {
                throw new InvalidDataException("The HORIZONTALPAGEBREAKS record declares too many page breaks.");
            }

            int expectedLength = checked(2 + (count * 6));
            if (expectedLength > payload.Length) {
                throw new InvalidDataException("The HORIZONTALPAGEBREAKS record ended before all breaks could be read.");
            }

            for (int i = 0; i < count; i++) {
                int breakOffset = 2 + (i * 6);
                ushort firstRowBelowBreak = BiffRecordReader.ReadUInt16(payload, breakOffset);
                ushort columnStart = BiffRecordReader.ReadUInt16(payload, breakOffset + 2);
                ushort columnEnd = BiffRecordReader.ReadUInt16(payload, breakOffset + 4);
                if (columnStart > 16383 || columnEnd > 16383 || columnEnd < columnStart) {
                    throw new InvalidDataException("The HORIZONTALPAGEBREAKS record contains an invalid column span.");
                }

                if (firstRowBelowBreak > 0) {
                    sheet.AddRowPageBreak(new LegacyXlsPageBreak(firstRowBelowBreak, columnStart + 1, columnEnd + 1));
                }
            }
        }

        private static void ParseMargin(LegacyXlsWorksheet sheet, byte[] payload, BiffRecordType type) {
            if (payload.Length < 8) {
                return;
            }

            double value = BiffRecordReader.ReadDouble(payload, 0);
            if (value < 0d || value > 49d) {
                return;
            }

            LegacyXlsPageSetup pageSetup = sheet.GetOrCreatePageSetup();
            switch (type) {
                case BiffRecordType.LeftMargin:
                    pageSetup.LeftMargin = value;
                    break;
                case BiffRecordType.RightMargin:
                    pageSetup.RightMargin = value;
                    break;
                case BiffRecordType.TopMargin:
                    pageSetup.TopMargin = value;
                    break;
                case BiffRecordType.BottomMargin:
                    pageSetup.BottomMargin = value;
                    break;
            }
        }

        private static void ParsePrintOption(LegacyXlsWorksheet sheet, byte[] payload, BiffRecordType type) {
            if (payload.Length < 2) {
                return;
            }

            bool value = (BiffRecordReader.ReadUInt16(payload, 0) & 0x0001) != 0;
            LegacyXlsPageSetup pageSetup = sheet.GetOrCreatePageSetup();
            switch (type) {
                case BiffRecordType.HCenter:
                    pageSetup.HorizontalCentered = value;
                    break;
                case BiffRecordType.PrintGrid:
                    pageSetup.PrintGridLines = value;
                    break;
                case BiffRecordType.PrintRowCol:
                    pageSetup.PrintHeadings = value;
                    break;
                case BiffRecordType.VCenter:
                    pageSetup.VerticalCentered = value;
                    break;
            }
        }

        private static void ParseUncalced(
            LegacyXlsWorksheet sheet,
            byte[] payload,
            int offset,
            ushort type,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            sheet.SetFullCalculationOnLoad(true);
            sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.Uncalced, offset, type);

            if (payload.Length < 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-WORKSHEET-UNCALCED-SHORT",
                    "The worksheet Uncalced record is shorter than expected.",
                    sheetName: sheet.Name,
                    recordOffset: offset,
                    recordType: type));
                return;
            }

            ushort reserved = BiffRecordReader.ReadUInt16(payload, 0);
            if (reserved != 0) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-WORKSHEET-UNCALCED-RESERVED",
                    $"The worksheet Uncalced record contains unexpected reserved value {reserved}.",
                    sheetName: sheet.Name,
                    recordOffset: offset,
                    recordType: type));
            }
        }

        private static bool TryParsePhoneticInfo(
            LegacyXlsWorksheet sheet,
            byte[] payload,
            int offset,
            ushort type,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out bool hasPhoneticRanges) {
            hasPhoneticRanges = false;
            if (payload.Length < 4) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-WORKSHEET-PHONETICINFO-SHORT",
                    "The worksheet PhoneticInfo record is shorter than expected.",
                    sheetName: sheet.Name,
                    recordOffset: offset,
                    recordType: type));
                return false;
            }

            ushort fontId = BiffRecordReader.ReadUInt16(payload, 0);
            ushort flags = BiffRecordReader.ReadUInt16(payload, 2);
            var phoneticType = (LegacyXlsPhoneticType)(flags & 0x0003);
            var phoneticAlignment = (LegacyXlsPhoneticAlignment)((flags >> 2) & 0x0003);

            int rangeOffset = 4;
            IReadOnlyList<string> ranges = Array.Empty<string>();
            bool rangePayloadUnsupported = false;
            if (rangeOffset < payload.Length) {
                rangePayloadUnsupported = !TryReadPhoneticRanges(payload, ref rangeOffset, out ranges)
                    || (rangeOffset != payload.Length && !HasOnlyZeroBytes(payload, rangeOffset));
            }

            sheet.SetPhoneticSettings(new LegacyXlsPhoneticSettings(fontId, phoneticType, phoneticAlignment, ranges));
            sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.PhoneticSettings, offset, type);
            hasPhoneticRanges = ranges.Count > 0 || rangePayloadUnsupported;
            return true;
        }

        private static bool HasOnlyZeroBytes(byte[] payload, int offset) {
            for (int i = offset; i < payload.Length; i++) {
                if (payload[i] != 0) {
                    return false;
                }
            }

            return true;
        }

        private static bool TryReadPhoneticRanges(byte[] payload, ref int offset, out IReadOnlyList<string> ranges) {
            ranges = Array.Empty<string>();
            if (offset + 2 > payload.Length) {
                return false;
            }

            ushort count = BiffRecordReader.ReadUInt16(payload, offset);
            offset += 2;
            if (count > 0x2000) {
                return false;
            }

            int expectedLength = checked(count * 8);
            if (offset + expectedLength > payload.Length) {
                return false;
            }

            if (count == 0) {
                return true;
            }

            var parsedRanges = new List<string>(count);
            for (int i = 0; i < count; i++) {
                ushort firstRow = BiffRecordReader.ReadUInt16(payload, offset);
                ushort lastRow = BiffRecordReader.ReadUInt16(payload, offset + 2);
                ushort firstColumn = BiffRecordReader.ReadUInt16(payload, offset + 4);
                ushort lastColumn = BiffRecordReader.ReadUInt16(payload, offset + 6);
                offset += 8;

                if (lastRow < firstRow || lastColumn < firstColumn || firstColumn > 0x00ff || lastColumn > 0x00ff) {
                    return false;
                }

                string start = A1.CellReference(firstRow + 1, firstColumn + 1);
                string end = A1.CellReference(lastRow + 1, lastColumn + 1);
                parsedRanges.Add(start == end ? start : start + ":" + end);
            }

            ranges = parsedRanges;
            return true;
        }

        private static void ParseVerticalPageBreaks(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 2) {
                return;
            }

            ushort count = BiffRecordReader.ReadUInt16(payload, 0);
            if (count > 255) {
                throw new InvalidDataException("The VERTICALPAGEBREAKS record declares too many page breaks.");
            }

            int expectedLength = checked(2 + (count * 6));
            if (expectedLength > payload.Length) {
                throw new InvalidDataException("The VERTICALPAGEBREAKS record ended before all breaks could be read.");
            }

            for (int i = 0; i < count; i++) {
                int breakOffset = 2 + (i * 6);
                ushort firstColumnRightOfBreak = BiffRecordReader.ReadUInt16(payload, breakOffset);
                ushort rowStart = BiffRecordReader.ReadUInt16(payload, breakOffset + 2);
                ushort rowEnd = BiffRecordReader.ReadUInt16(payload, breakOffset + 4);
                if (rowEnd < rowStart) {
                    throw new InvalidDataException("The VERTICALPAGEBREAKS record contains an invalid row span.");
                }

                if (firstColumnRightOfBreak > 0) {
                    sheet.AddColumnPageBreak(new LegacyXlsPageBreak(firstColumnRightOfBreak, rowStart + 1, rowEnd + 1));
                }
            }
        }

        private static void ParseFormula(
            LegacyXlsWorksheet sheet,
            byte[] payload,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            List<LegacyXlsFormulaTokenRecord> formulaTokenRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options,
            BiffSharedFormulaImportState sharedFormulaState,
            int recordOffset,
            ref PendingFormulaString? pendingFormulaString) {
            if (payload.Length < 20) {
                return;
            }

            int row = BiffRecordReader.ReadUInt16(payload, 0) + 1;
            int column = BiffRecordReader.ReadUInt16(payload, 2) + 1;
            ushort styleIndex = BiffRecordReader.ReadUInt16(payload, 4);
            BiffFormulaValue formulaValue = BiffFormulaValueReader.Read(payload, 6);
            BiffSharedFormulaReference? sharedFormulaReference = null;
            bool isSharedFormulaReference = BiffSharedFormulaImportState.TryReadFormulaReference(payload, 20, out BiffSharedFormulaReference formulaReference);
            if (isSharedFormulaReference) {
                sharedFormulaReference = formulaReference;
            }

            if (HasFormulaTokenPayload(payload)) {
                BiffFormulaTokenScanner.ScanLengthPrefixed(
                    payload,
                    20,
                    isSharedFormulaReference ? "SharedFormulaReference" : "CellFormula",
                    sheet.Name,
                    A1.ColumnIndexToLetters(column) + row.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    recordOffset,
                    (ushort)BiffRecordType.Formula,
                    formulaTokenRecords);
            }

            string? formulaText = null;
            BiffFormulaReadFailure? formulaFailure = null;
            bool formulaTextRead = !isSharedFormulaReference
                && BiffFormulaTextReader.TryRead(payload, 20, row - 1, column - 1, externSheets, externalReferences, sheetNames, definedNames, out formulaText, out formulaFailure);
            if (!isSharedFormulaReference && !formulaTextRead && options.ReportUnsupportedContent && HasFormulaTokenPayload(payload)) {
                string failureDescription = formulaFailure == null ? "Unsupported formula tokens" : formulaFailure.Description;
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Info,
                    "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED",
                    $"{failureDescription} Formula at {A1.ColumnIndexToLetters(column)}{row} was imported from its cached result.",
                    sheetName: sheet.Name,
                    recordOffset: recordOffset,
                    recordType: (ushort)BiffRecordType.Formula,
                    detailCode: formulaFailure?.DetailCode,
                    formulaContext: "CellFormula",
                    formulaToken: formulaFailure?.Token,
                    formulaTokenName: formulaFailure?.TokenName,
                    formulaTokenOffset: formulaFailure?.TokenOffset));
            }

            if (formulaValue.ExpectsStringRecord) {
                pendingFormulaString = new PendingFormulaString(row, column, styleIndex, formulaText, sharedFormulaReference, recordOffset);
                return;
            }

            sheet.AddCell(new LegacyXlsCell(
                row,
                column,
                formulaValue.Kind,
                formulaValue.Value,
                styleIndex,
                isFormula: true,
                formulaText: formulaText));
            sharedFormulaState.RegisterFormulaCell(row, column, sharedFormulaReference, recordOffset);
        }

        private static bool HasFormulaTokenPayload(byte[] payload) {
            if (payload.Length < 22) {
                return false;
            }

            ushort expressionLength = BiffRecordReader.ReadUInt16(payload, 20);
            return expressionLength > 0 && 22 + expressionLength <= payload.Length;
        }

        private static void FlushPendingFormulaString(
            LegacyXlsWorksheet sheet,
            BiffSharedFormulaImportState sharedFormulaState,
            ref PendingFormulaString? pendingFormulaString,
            string value = "") {
            if (pendingFormulaString == null) {
                return;
            }

            PendingFormulaString formulaString = pendingFormulaString;
            sheet.AddCell(formulaString.ToCell(formulaString.GetValue(value)));
            sharedFormulaState.RegisterFormulaCell(formulaString.Row, formulaString.Column, formulaString.SharedFormulaReference, formulaString.RecordOffset);
            pendingFormulaString = null;
        }

        private static void ParseColInfo(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 12) {
                return;
            }

            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, 0);
            ushort lastColumn = BiffRecordReader.ReadUInt16(payload, 2);
            if (lastColumn < firstColumn) {
                throw new InvalidDataException("The COLINFO record has an invalid column range.");
            }

            ushort widthUnits = BiffRecordReader.ReadUInt16(payload, 4);
            ushort styleIndex = BiffRecordReader.ReadUInt16(payload, 6);
            ushort options = BiffRecordReader.ReadUInt16(payload, 8);
            byte outlineLevel = (byte)((options >> 8) & 0x0007);
            bool collapsed = (options & 0x1000) != 0;
            sheet.AddColumn(new LegacyXlsColumnLayout(
                firstColumn + 1,
                lastColumn + 1,
                widthUnits / 256d,
                (options & 0x0001) != 0,
                styleIndex,
                outlineLevel,
                collapsed));
        }

        private static void ParseFrozenPane(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 10) {
                return;
            }

            ushort leftColumns = BiffRecordReader.ReadUInt16(payload, 0);
            ushort topRows = BiffRecordReader.ReadUInt16(payload, 2);
            if (leftColumns == 0 && topRows == 0) {
                return;
            }

            sheet.SetFreezePane(new LegacyXlsFreezePane(topRows, leftColumns));
        }

        private static bool TryParseSplitPane(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 10 || !HasSplitPane(payload)) {
                return false;
            }

            ushort horizontalSplit = BiffRecordReader.ReadUInt16(payload, 0);
            ushort verticalSplit = BiffRecordReader.ReadUInt16(payload, 2);
            ushort topRow = BiffRecordReader.ReadUInt16(payload, 4);
            ushort leftColumn = BiffRecordReader.ReadUInt16(payload, 6);
            byte activePane = payload[8];
            if (activePane > 3) {
                return false;
            }

            sheet.SetSplitPane(new LegacyXlsSplitPane(horizontalSplit, verticalSplit, topRow, leftColumn, activePane));
            return true;
        }

        private static bool HasSplitPane(byte[] payload) {
            if (payload.Length < 4) {
                return false;
            }

            return BiffRecordReader.ReadUInt16(payload, 0) != 0
                || BiffRecordReader.ReadUInt16(payload, 2) != 0;
        }

        private static void ParseHeaderFooter(LegacyXlsWorksheet sheet, byte[] payload, bool isHeader) {
            if (payload.Length == 0) {
                return;
            }

            int offset = 0;
            string text = BiffStringReader.ReadUnicodeString(payload, ref offset);
            if (text.Length == 0) {
                return;
            }

            LegacyXlsPageSetup pageSetup = sheet.GetOrCreatePageSetup();
            if (isHeader) {
                pageSetup.HeaderText = text;
            } else {
                pageSetup.FooterText = text;
            }
        }

        private static bool TryParseHeaderFooterExtension(LegacyXlsWorksheet sheet, byte[] payload, int recordOffset) {
            const int MinimumPayloadLength = 38;
            if (payload.Length < MinimumPayloadLength) {
                return false;
            }

            ushort wrappedRecordType = BiffRecordReader.ReadUInt16(payload, 0);
            if (wrappedRecordType != (ushort)BiffRecordType.HeaderFooter) {
                return false;
            }

            ushort flags = BiffRecordReader.ReadUInt16(payload, 28);
            int offset = 30;
            ushort evenHeaderLength = BiffRecordReader.ReadUInt16(payload, offset);
            offset += 2;
            ushort evenFooterLength = BiffRecordReader.ReadUInt16(payload, offset);
            offset += 2;
            ushort firstHeaderLength = BiffRecordReader.ReadUInt16(payload, offset);
            offset += 2;
            ushort firstFooterLength = BiffRecordReader.ReadUInt16(payload, offset);
            offset += 2;

            try {
                string evenHeader = ReadHeaderFooterExtensionString(payload, ref offset, evenHeaderLength);
                string evenFooter = ReadHeaderFooterExtensionString(payload, ref offset, evenFooterLength);
                string firstHeader = ReadHeaderFooterExtensionString(payload, ref offset, firstHeaderLength);
                string firstFooter = ReadHeaderFooterExtensionString(payload, ref offset, firstFooterLength);

                LegacyXlsPageSetup pageSetup = sheet.GetOrCreatePageSetup();
                pageSetup.DifferentOddEvenHeaderFooter = (flags & 0x0001) != 0;
                pageSetup.DifferentFirstHeaderFooter = (flags & 0x0002) != 0;
                pageSetup.ScaleHeaderFooterWithDocument = (flags & 0x0004) != 0;
                pageSetup.AlignHeaderFooterWithMargins = (flags & 0x0008) != 0;
                pageSetup.EvenHeaderText = evenHeader.Length == 0 ? null : evenHeader;
                pageSetup.EvenFooterText = evenFooter.Length == 0 ? null : evenFooter;
                pageSetup.FirstHeaderText = firstHeader.Length == 0 ? null : firstHeader;
                pageSetup.FirstFooterText = firstFooter.Length == 0 ? null : firstFooter;
                sheet.AddMetadataRecord(LegacyXlsWorksheetMetadataKind.HeaderFooter, recordOffset, (ushort)BiffRecordType.HeaderFooter);
                return true;
            } catch (InvalidDataException) {
                return false;
            }
        }

        private static string ReadHeaderFooterExtensionString(byte[] payload, ref int offset, int charCount) {
            if (charCount == 0) {
                return string.Empty;
            }

            return BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, charCount);
        }

        private static void ParseMergeCells(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 2) {
                return;
            }

            ushort count = BiffRecordReader.ReadUInt16(payload, 0);
            int expectedLength = checked(2 + (count * 8));
            if (expectedLength > payload.Length) {
                throw new InvalidDataException("The MERGECELLS record ended before all ranges could be read.");
            }

            for (int i = 0; i < count; i++) {
                int rangeOffset = 2 + (i * 8);
                ushort firstRow = BiffRecordReader.ReadUInt16(payload, rangeOffset);
                ushort lastRow = BiffRecordReader.ReadUInt16(payload, rangeOffset + 2);
                ushort firstColumn = BiffRecordReader.ReadUInt16(payload, rangeOffset + 4);
                ushort lastColumn = BiffRecordReader.ReadUInt16(payload, rangeOffset + 6);
                if (lastRow < firstRow || lastColumn < firstColumn) {
                    throw new InvalidDataException("The MERGECELLS record contains an invalid range.");
                }

                sheet.AddMergedRange(new LegacyXlsMergedRange(
                    firstRow + 1,
                    firstColumn + 1,
                    lastRow + 1,
                    lastColumn + 1));
            }
        }

        private static void ParsePassword(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 2) {
                return;
            }

            ushort passwordHash = BiffRecordReader.ReadUInt16(payload, 0);
            if (passwordHash != 0) {
                sheet.SetProtectionPasswordHash(passwordHash);
            }
        }

        private static void ParseProtect(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 2) {
                return;
            }

            sheet.SetProtection(BiffRecordReader.ReadUInt16(payload, 0) != 0);
        }

        private static void ParseObjectProtection(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 2) {
                return;
            }

            sheet.SetObjectProtection(BiffRecordReader.ReadUInt16(payload, 0) != 0);
        }

        private static void ParseScenarioProtection(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 2) {
                return;
            }

            sheet.SetScenarioProtection(BiffRecordReader.ReadUInt16(payload, 0) != 0);
        }

        private static void ParseSortSettings(
            LegacyXlsWorksheet sheet,
            byte[] payload,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options,
            int recordOffset) {
            if (payload.Length < 6) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-SORT-SHORT",
                    "The Sort record is too short.",
                    sheetName: sheet.Name,
                    recordOffset: recordOffset,
                    recordType: (ushort)BiffRecordType.Sort));
                return;
            }

            ushort flags = BiffRecordReader.ReadUInt16(payload, 0);
            int offset = 5;
            string? key1 = ReadSortKey(payload, ref offset, payload[2]);
            string? key2 = ReadSortKey(payload, ref offset, payload[3]);
            string? key3 = ReadSortKey(payload, ref offset, payload[4]);
            int customListIndex = (flags >> 5) & 0x001f;
            if (offset >= payload.Length) {
                throw new InvalidDataException("The Sort record ended before the reserved byte.");
            }

            byte reserved = payload[offset++];
            if ((flags & 0xf800) != 0 || reserved != 0 || offset != payload.Length) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Info,
                    "XLS-BIFF-SORT-RESERVED-BYTES",
                    "The Sort record contains non-default reserved data; parsed known sort settings only.",
                    sheetName: sheet.Name,
                    recordOffset: recordOffset,
                    recordType: (ushort)BiffRecordType.Sort));
            }

            if (customListIndex != 0) {
                AddUnsupportedSortCustomListFeature(sheet, unsupportedFeatures, diagnostics, options, recordOffset, customListIndex);
            }

            sheet.SetSortSettings(new LegacyXlsSortSettings(
                (flags & 0x0001) != 0,
                (flags & 0x0002) != 0,
                (flags & 0x0004) != 0,
                (flags & 0x0008) != 0,
                (flags & 0x0010) != 0,
                customListIndex,
                (flags & 0x0400) != 0,
                key1,
                key2,
                key3));
        }

        private static void AddUnsupportedSortCustomListFeature(
            LegacyXlsWorksheet sheet,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options,
            int recordOffset,
            int customListIndex) {
            string description = $"The worksheet Sort record uses custom-list sort order index {customListIndex}; sort keys are projected, but the environment-dependent custom-list order cannot be projected to an Open XML custom-list string.";
            var feature = new LegacyXlsUnsupportedFeature(
                LegacyXlsUnsupportedFeatureKind.WorksheetSort,
                "XLS-BIFF-FEATURE-WORKSHEET-SORT-CUSTOM-LIST-UNSUPPORTED",
                description,
                sheet.Name,
                recordOffset,
                (ushort)BiffRecordType.Sort,
                "WorksheetSort:CustomListIndex");
            unsupportedFeatures.Add(feature);

            if (options.ReportUnsupportedContent) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    feature.Code,
                    description,
                    sheetName: sheet.Name,
                    recordOffset: recordOffset,
                    recordType: (ushort)BiffRecordType.Sort));
            }
        }

        private static string? ReadSortKey(byte[] payload, ref int offset, byte characterCount) {
            if (characterCount == 0) {
                return null;
            }

            return BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, characterCount);
        }

        private static void ParseSetup(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 34) {
                return;
            }

            LegacyXlsPageSetup pageSetup = sheet.GetOrCreatePageSetup();
            ushort scale = BiffRecordReader.ReadUInt16(payload, 2);
            ushort fitToWidth = BiffRecordReader.ReadUInt16(payload, 6);
            ushort fitToHeight = BiffRecordReader.ReadUInt16(payload, 8);
            ushort flags = BiffRecordReader.ReadUInt16(payload, 10);
            bool ignorePrinterSettings = (flags & 0x0004) != 0;
            bool ignoreOrientation = (flags & 0x0040) != 0;

            if (!ignorePrinterSettings && scale >= 10 && scale <= 400) {
                pageSetup.Scale = scale;
            }

            pageSetup.FitToWidth = fitToWidth;
            pageSetup.FitToHeight = fitToHeight;
            pageSetup.PageOrder = (flags & 0x0001) != 0
                ? ExcelPageOrder.OverThenDown
                : ExcelPageOrder.DownThenOver;

            if (!ignorePrinterSettings && !ignoreOrientation) {
                pageSetup.Landscape = (flags & 0x0002) == 0;
            }

            double header = BiffRecordReader.ReadDouble(payload, 16);
            if (header >= 0d && header < 49d) {
                pageSetup.HeaderMargin = header;
            }

            double footer = BiffRecordReader.ReadDouble(payload, 24);
            if (footer >= 0d && footer < 49d) {
                pageSetup.FooterMargin = footer;
            }
        }

        private static void ParseMulBlank(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 6) {
                return;
            }

            ushort row = BiffRecordReader.ReadUInt16(payload, 0);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, 2);
            ushort lastColumn = BiffRecordReader.ReadUInt16(payload, payload.Length - 2);
            int cellCount = checked(lastColumn - firstColumn + 1);
            int expectedLength = checked(4 + (cellCount * 2) + 2);
            if (cellCount < 1 || expectedLength > payload.Length) {
                throw new InvalidDataException("The MULBLANK record has an invalid column range.");
            }

            for (int i = 0; i < cellCount; i++) {
                ushort styleIndex = BiffRecordReader.ReadUInt16(payload, 4 + (i * 2));
                sheet.AddCell(new LegacyXlsCell(
                    row + 1,
                    firstColumn + i + 1,
                    LegacyXlsCellValueKind.Blank,
                    null,
                    styleIndex));
            }
        }

        private static void ParseMulRk(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 10) {
                return;
            }

            ushort row = BiffRecordReader.ReadUInt16(payload, 0);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, 2);
            ushort lastColumn = BiffRecordReader.ReadUInt16(payload, payload.Length - 2);
            int cellCount = checked(lastColumn - firstColumn + 1);
            int expectedLength = checked(4 + (cellCount * 6) + 2);
            if (cellCount < 1 || expectedLength > payload.Length) {
                throw new InvalidDataException("The MULRK record has an invalid column range.");
            }

            for (int i = 0; i < cellCount; i++) {
                int cellOffset = 4 + (i * 6);
                sheet.AddCell(new LegacyXlsCell(
                    row + 1,
                    firstColumn + i + 1,
                    LegacyXlsCellValueKind.Number,
                    BiffRkNumberReader.ReadRkNumber(BiffRecordReader.ReadUInt32(payload, cellOffset + 2)),
                    BiffRecordReader.ReadUInt16(payload, cellOffset)));
            }
        }

        private static void ParseRow(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 16) {
                return;
            }

            ushort row = BiffRecordReader.ReadUInt16(payload, 0);
            ushort heightTwips = BiffRecordReader.ReadUInt16(payload, 6);
            ushort options = BiffRecordReader.ReadUInt16(payload, 12);
            byte outlineLevel = (byte)(options & 0x0007);
            bool collapsed = (options & 0x0010) != 0;
            bool hidden = (options & 0x0020) != 0;
            bool customHeight = (options & 0x0040) != 0;
            bool customFormat = (options & 0x0080) != 0;
            ushort? styleIndex = customFormat
                ? (ushort)(BiffRecordReader.ReadUInt16(payload, 14) & 0x0fff)
                : null;
            double heightPoints = heightTwips / 20d;
            if (!hidden && !customHeight && !customFormat && outlineLevel == 0 && !collapsed && heightPoints <= 0) {
                return;
            }

            sheet.AddRow(new LegacyXlsRowLayout(row + 1, heightPoints, hidden, customHeight, styleIndex, outlineLevel, collapsed));
        }

        private static void ParseDefaultColumnWidth(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 2) {
                return;
            }

            ushort width = BiffRecordReader.ReadUInt16(payload, 0);
            if (width == 0 || width > 255) {
                return;
            }

            sheet.SetDefaultColumnWidth(width);
        }

        private static void ParseDefaultRowHeight(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 4) {
                return;
            }

            ushort options = BiffRecordReader.ReadUInt16(payload, 0);
            bool hidden = (options & 0x0002) != 0;
            short heightTwips = BiffRecordReader.ReadInt16(payload, 2);
            if (heightTwips <= 0 || heightTwips > 8179) {
                return;
            }

            sheet.SetDefaultRowHeight(heightTwips / 20d, hidden);
        }

        private static void ParseWindow2(LegacyXlsWorksheet sheet, byte[] payload, out bool frozenWindow) {
            frozenWindow = false;
            if (payload.Length < 2) {
                return;
            }

            ushort options = BiffRecordReader.ReadUInt16(payload, 0);
            frozenWindow = (options & 0x0008) != 0;
            bool showFormulas = (options & 0x0001) != 0;
            bool showGridLines = (options & 0x0002) != 0;
            bool showRowColumnHeadings = (options & 0x0004) != 0;
            bool showZeroValues = (options & 0x0010) != 0;
            bool defaultGridColor = (options & 0x0020) != 0;
            bool rightToLeft = (options & 0x0040) != 0;
            bool showOutlineSymbols = (options & 0x0080) != 0;
            bool frozenWithoutSplit = frozenWindow && (options & 0x0100) != 0;
            bool tabSelected = (options & 0x0200) != 0;
            bool pageBreakPreview = (options & 0x0800) != 0;

            sheet.SetFormulasVisible(showFormulas);
            sheet.SetGridLinesVisible(showGridLines);
            sheet.SetRowColumnHeadingsVisible(showRowColumnHeadings);
            sheet.SetZeroValuesVisible(showZeroValues);
            sheet.SetDefaultGridColor(defaultGridColor);
            sheet.SetRightToLeft(rightToLeft);
            sheet.SetOutlineSymbolsVisible(showOutlineSymbols);
            sheet.SetFrozenWithoutSplit(frozenWithoutSplit);
            sheet.SetTabSelected(tabSelected);
            sheet.SetPageBreakPreview(pageBreakPreview);

            ushort? gridLineColorIndex = null;
            if (payload.Length >= 8) {
                ushort parsedGridLineColorIndex = BiffRecordReader.ReadUInt16(payload, 6);
                if (parsedGridLineColorIndex <= 64) {
                    gridLineColorIndex = parsedGridLineColorIndex;
                    sheet.SetGridLineColorIndex(parsedGridLineColorIndex);
                }
            }

            int? firstVisibleRow = null;
            int? firstVisibleColumn = null;
            if (payload.Length >= 6) {
                ushort parsedFirstVisibleRow = BiffRecordReader.ReadUInt16(payload, 2);
                ushort parsedFirstVisibleColumn = BiffRecordReader.ReadUInt16(payload, 4);
                if (parsedFirstVisibleRow != ushort.MaxValue && parsedFirstVisibleColumn != byte.MaxValue) {
                    firstVisibleRow = parsedFirstVisibleRow;
                    firstVisibleColumn = parsedFirstVisibleColumn;
                    sheet.SetFirstVisibleCell(parsedFirstVisibleRow, parsedFirstVisibleColumn);
                }
            }

            uint? zoomScale = null;
            uint? zoomScaleNormal = null;
            if (payload.Length >= 14) {
                ushort pageBreakPreviewZoom = BiffRecordReader.ReadUInt16(payload, 10);
                if (pageBreakPreview && IsValidWindow2Zoom(pageBreakPreviewZoom)) {
                    zoomScale = pageBreakPreviewZoom;
                    sheet.SetZoomScale(pageBreakPreviewZoom);
                }

                ushort normalZoom = BiffRecordReader.ReadUInt16(payload, 12);
                if (IsValidWindow2Zoom(normalZoom)) {
                    zoomScaleNormal = normalZoom;
                    sheet.SetZoomScaleNormal(normalZoom);
                }
            }

            sheet.AddWindowView(new LegacyXlsWorksheetWindowView(
                showFormulas,
                showGridLines,
                showRowColumnHeadings,
                showZeroValues,
                rightToLeft,
                defaultGridColor,
                gridLineColorIndex,
                showOutlineSymbols,
                tabSelected,
                pageBreakPreview,
                frozenWithoutSplit,
                firstVisibleRow,
                firstVisibleColumn,
                zoomScale,
                zoomScaleNormal));
        }

        private static bool IsValidWindow2Zoom(ushort zoomScale) {
            return zoomScale >= 10 && zoomScale <= 400;
        }

        private static bool TryParseSheetExtension(LegacyXlsWorksheet sheet, byte[] payload, out bool hasUnsupportedSheetExtensionMetadata) {
            hasUnsupportedSheetExtensionMetadata = false;
            if (payload.Length < 20) {
                return false;
            }

            ushort wrappedRecordType = BiffRecordReader.ReadUInt16(payload, 0);
            if (wrappedRecordType != (ushort)BiffRecordType.SheetExt) {
                return false;
            }

            uint declaredByteCount = BiffRecordReader.ReadUInt32(payload, 12);
            if (declaredByteCount < 20U || declaredByteCount > payload.Length) {
                return false;
            }

            ushort colorIndex = (ushort)(BiffRecordReader.ReadUInt32(payload, 16) & 0x7FU);
            if (colorIndex != 0x7F) {
                sheet.SetTabColorIndex(colorIndex);
            }

            hasUnsupportedSheetExtensionMetadata = declaredByteCount > 20U;
            return true;
        }

        private sealed class PendingFormulaString {
            private List<byte[]>? _stringPayloads;

            internal PendingFormulaString(
                int row,
                int column,
                ushort styleIndex,
                string? formulaText,
                BiffSharedFormulaReference? sharedFormulaReference,
                int recordOffset) {
                Row = row;
                Column = column;
                StyleIndex = styleIndex;
                FormulaText = formulaText;
                SharedFormulaReference = sharedFormulaReference;
                RecordOffset = recordOffset;
            }

            internal int Row { get; }

            internal int Column { get; }

            private ushort StyleIndex { get; }

            private string? FormulaText { get; }

            internal BiffSharedFormulaReference? SharedFormulaReference { get; }

            internal int RecordOffset { get; }

            internal bool HasStringPayload => _stringPayloads?.Count > 0;

            internal void AddStringPayload(byte[] payload) {
                _stringPayloads ??= new List<byte[]>();
                _stringPayloads.Add(payload);
            }

            internal string GetValue(string fallback) {
                return _stringPayloads == null || _stringPayloads.Count == 0
                    ? fallback
                    : BiffStringReader.ReadUnicodeString(_stringPayloads);
            }

            internal LegacyXlsCell ToCell(string value) {
                return new LegacyXlsCell(
                    Row,
                    Column,
                    LegacyXlsCellValueKind.Text,
                    value,
                    StyleIndex,
                    isFormula: true,
                    formulaText: FormulaText);
            }
        }
    }
}
