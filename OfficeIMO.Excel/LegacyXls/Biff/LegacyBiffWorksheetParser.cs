using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class LegacyBiffWorksheetParser {
        internal static void Parse(
            byte[] workbookStream,
            LegacyXlsWorksheet sheet,
            IReadOnlyList<string> sharedStrings,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options) {
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
            PendingFormulaString? pendingFormulaString = null;
            var commentState = new BiffCommentImportState(sheet);
            var conditionalFormattingState = new BiffConditionalFormattingImportState(sheet, externSheets, externalReferences, sheetNames, definedNames);
            var sharedFormulaState = new BiffSharedFormulaImportState(sheet, externSheets, externalReferences, sheetNames, definedNames, diagnostics, options);
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

                    if (!LegacyBiffVersionValidator.ValidateWorksheetBof(payload, offset, sheet.Name, unsupportedFeatures, diagnostics)) {
                        return;
                    }
                }

                if (type == (ushort)BiffRecordType.Eof) {
                    FlushPendingFormulaString(sheet, sharedFormulaState, ref pendingFormulaString);
                    commentState.AddUnresolvedFeatures(unsupportedFeatures, diagnostics, options.ReportUnsupportedRecords);
                    conditionalFormattingState.AddUnresolvedFeatures(unsupportedFeatures, diagnostics, options.ReportUnsupportedRecords);
                    sharedFormulaState.AddUnresolvedDiagnostics();
                    return;
                }

                ParseWorksheetRecord(sheet, sharedStrings, externSheets, externalReferences, sheetNames, definedNames, unsupportedFeatures, diagnostics, options, commentState, conditionalFormattingState, sharedFormulaState, type, offset, payload, ref frozenWindow, ref pendingFormulaString);
                offset = payloadOffset + length;
            }

            FlushPendingFormulaString(sheet, sharedFormulaState, ref pendingFormulaString);
            commentState.AddUnresolvedFeatures(unsupportedFeatures, diagnostics, options.ReportUnsupportedRecords);
            conditionalFormattingState.AddUnresolvedFeatures(unsupportedFeatures, diagnostics, options.ReportUnsupportedRecords);
            sharedFormulaState.AddUnresolvedDiagnostics();
        }

        private static void ParseWorksheetRecord(
            LegacyXlsWorksheet sheet,
            IReadOnlyList<string> sharedStrings,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options,
            BiffCommentImportState commentState,
            BiffConditionalFormattingImportState conditionalFormattingState,
            BiffSharedFormulaImportState sharedFormulaState,
            ushort type,
            int offset,
            byte[] payload,
            ref bool frozenWindow,
            ref PendingFormulaString? pendingFormulaString) {
            try {
                if (pendingFormulaString != null && type != (ushort)BiffRecordType.String) {
                    FlushPendingFormulaString(sheet, sharedFormulaState, ref pendingFormulaString);
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
                    case BiffRecordType.CondFmt:
                        if (!conditionalFormattingState.TryReadHeader(payload, offset)) {
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
                        }

                        break;
                    case BiffRecordType.Cf:
                        if (!conditionalFormattingState.TryReadRule(payload)) {
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
                        }

                        break;
                    case BiffRecordType.Continue:
                        commentState.TryReadContinue(payload);
                        break;
                    case BiffRecordType.DefColWidth:
                        ParseDefaultColumnWidth(sheet, payload);
                        break;
                    case BiffRecordType.AutoFilterInfo:
                        if (BiffAutoFilterReader.TryReadInfo(payload, out ushort dropDownCount)) {
                            sheet.SetAutoFilterDropDownCount(dropDownCount);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
                        }

                        break;
                    case BiffRecordType.AutoFilter:
                        if (BiffAutoFilterReader.TryReadCriteria(payload, out LegacyXlsAutoFilterCriteria? criteria)) {
                            sheet.AddAutoFilterCriteria(criteria!);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
                        }

                        break;
                    case BiffRecordType.FilterMode:
                        break;
                    case BiffRecordType.DefaultRowHeight:
                        ParseDefaultRowHeight(sheet, payload);
                        break;
                    case BiffRecordType.Dimensions:
                        ParseDimensions(sheet, payload);
                        break;
                    case BiffRecordType.DVal:
                        if (!IsValidDataValidationCollectionHeader(payload)) {
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
                        }

                        break;
                    case BiffRecordType.Dv:
                        if (BiffDataValidationReader.TryRead(payload, externSheets, externalReferences, sheetNames, definedNames, out LegacyXlsDataValidation? validation)) {
                            sheet.AddDataValidation(validation!);
                        } else {
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
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
                    case BiffRecordType.Formula:
                        ParseFormula(sheet, payload, externSheets, externalReferences, sheetNames, definedNames, diagnostics, options, sharedFormulaState, offset, ref pendingFormulaString);
                        break;
                    case BiffRecordType.Footer:
                        ParseHeaderFooter(sheet, payload, isHeader: false);
                        break;
                    case BiffRecordType.Header:
                        ParseHeaderFooter(sheet, payload, isHeader: true);
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
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
                        }

                        break;
                    case BiffRecordType.Label:
                        if (payload.Length >= 8) {
                            int stringOffset = 6;
                            sheet.AddCell(new LegacyXlsCell(
                                BiffRecordReader.ReadUInt16(payload, 0) + 1,
                                BiffRecordReader.ReadUInt16(payload, 2) + 1,
                                LegacyXlsCellValueKind.Text,
                                BiffStringReader.ReadUnicodeString(payload, ref stringOffset),
                                BiffRecordReader.ReadUInt16(payload, 4)));
                        }

                        break;
                    case BiffRecordType.LabelSst:
                        if (payload.Length >= 10) {
                            uint sharedStringIndex = BiffRecordReader.ReadUInt32(payload, 6);
                            string text = sharedStringIndex < sharedStrings.Count ? sharedStrings[(int)sharedStringIndex] : string.Empty;
                            sheet.AddCell(new LegacyXlsCell(
                                BiffRecordReader.ReadUInt16(payload, 0) + 1,
                                BiffRecordReader.ReadUInt16(payload, 2) + 1,
                                LegacyXlsCellValueKind.Text,
                                text,
                                BiffRecordReader.ReadUInt16(payload, 4)));
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
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
                        }

                        break;
                    case BiffRecordType.Obj:
                        if (!commentState.TryReadObject(payload)) {
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
                        }

                        break;
                    case BiffRecordType.Pane:
                        if (frozenWindow) {
                            ParseFrozenPane(sheet, payload);
                        }

                        break;
                    case BiffRecordType.Password:
                        ParsePassword(sheet, payload);
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
                    case BiffRecordType.Setup:
                        ParseSetup(sheet, payload);
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
                            int stringOffset = 0;
                            FlushPendingFormulaString(sheet, sharedFormulaState, ref pendingFormulaString, BiffStringReader.ReadUnicodeString(payload, ref stringOffset));
                        } else if (options.ReportUnsupportedRecords) {
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
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
                        }

                        break;
                    case BiffRecordType.TopMargin:
                        ParseMargin(sheet, payload, BiffRecordType.TopMargin);
                        break;
                    case BiffRecordType.Txo:
                        if (!commentState.TryReadTextObject(payload)) {
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
                        }

                        break;
                    case BiffRecordType.VerticalPageBreaks:
                        ParseVerticalPageBreaks(sheet, payload);
                        break;
                    case BiffRecordType.Window2:
                        ParseWindow2(sheet, payload, out frozenWindow);
                        break;
                    default:
                        if (type != (ushort)BiffRecordType.Bof) {
                            AddUnsupportedFeature(unsupportedFeatures, diagnostics, options, type, offset, sheet.Name);
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

        private static bool IsValidDataValidationCollectionHeader(byte[] payload) {
            if (payload.Length < 18) {
                return false;
            }

            uint validationCount = BiffRecordReader.ReadUInt32(payload, 14);
            return validationCount <= 65534;
        }

        private static void AddUnsupportedFeature(
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsImportDiagnostic> diagnostics,
            LegacyXlsImportOptions options,
            ushort type,
            int offset,
            string? sheetName) {
            unsupportedFeatures.Add(BiffUnsupportedRecordDiagnostics.CreateUnsupportedRecordFeature(type, offset, sheetName));
            if (options.ReportUnsupportedRecords) {
                BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(diagnostics, type, offset, sheetName);
            }
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

        private static void ParseDimensions(LegacyXlsWorksheet sheet, byte[] payload) {
            if (payload.Length < 14) {
                return;
            }

            uint firstRow = BiffRecordReader.ReadUInt32(payload, 0);
            uint rowAfterLast = BiffRecordReader.ReadUInt32(payload, 4);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, 8);
            ushort columnAfterLast = BiffRecordReader.ReadUInt16(payload, 10);

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
                if (columnStart > 16383 || columnEnd > 16383 || columnEnd <= columnStart) {
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
                if (rowEnd <= rowStart) {
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

            string? formulaText = null;
            BiffFormulaReadFailure? formulaFailure = null;
            bool formulaTextRead = !isSharedFormulaReference
                && BiffFormulaTextReader.TryRead(payload, 20, row - 1, column - 1, externSheets, externalReferences, sheetNames, definedNames, out formulaText, out formulaFailure);
            if (!isSharedFormulaReference && !formulaTextRead && options.ReportUnsupportedRecords && HasFormulaTokenPayload(payload)) {
                string failureDescription = formulaFailure == null ? "Unsupported formula tokens" : formulaFailure.Description;
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Info,
                    "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED",
                    $"{failureDescription} Formula at {A1.ColumnIndexToLetters(column)}{row} was imported from its cached result.",
                    sheetName: sheet.Name,
                    recordOffset: recordOffset,
                    recordType: (ushort)BiffRecordType.Formula,
                    detailCode: formulaFailure?.DetailCode));
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
            sheet.AddCell(formulaString.ToCell(value));
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

            if (!ignoreOrientation) {
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
            sheet.SetGridLinesVisible((options & 0x0002) != 0);
            sheet.SetRowColumnHeadingsVisible((options & 0x0004) != 0);
            sheet.SetZeroValuesVisible((options & 0x0010) != 0);
            sheet.SetRightToLeft((options & 0x0040) != 0);
        }

        private sealed class PendingFormulaString {
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
