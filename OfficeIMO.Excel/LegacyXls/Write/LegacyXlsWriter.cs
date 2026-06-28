using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private static readonly byte[] WorkbookGlobalsBof = { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 };
        private static readonly byte[] WorksheetBof = { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 };

        internal static byte[] WriteWorkbook(ExcelDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            byte[] workbookStream = BuildWorkbookStream(document);
            IReadOnlyList<LegacyXlsCompoundStream> propertyStreams = LegacyOlePropertySetWriter.CreateDocumentPropertyStreams(document);
            return LegacyXlsCompoundFileWriter.Write(workbookStream, propertyStreams);
        }

        private static byte[] BuildWorkbookStream(ExcelDocument document) {
            var sheets = document.Sheets.ToArray();
            if (sheets.Length == 0) {
                throw new InvalidOperationException("Native XLS saving requires at least one worksheet.");
            }

            LegacyXlsFontTable fontTable = LegacyXlsFontTable.Create(document);
            LegacyXlsExternSheetTable externSheetTable = LegacyXlsExternSheetTable.Create(document, sheets);
            LegacyXlsFormulaNameIndex formulaNameIndex = LegacyXlsDefinedNameWriter.CreateFormulaNameIndex(document, sheets, externSheetTable);
            LegacyXlsWritePreflight.ThrowIfUnsupported(document, sheets, fontTable, formulaNameIndex);
            LegacyXlsStyleTable styleTable = LegacyXlsStyleTable.Create(document, sheets, fontTable);
            ReserveWorksheetTabColors(sheets, fontTable);

            using var stream = new MemoryStream();
            WriteRecord(stream, 0x0809, WorkbookGlobalsBof);
            WriteRecord(stream, 0x0042, BuildUInt16Payload(1200));
            string? workbookCodeName = document.WorkbookRoot.GetFirstChild<WorkbookProperties>()?.CodeName?.Value;
            if (!string.IsNullOrWhiteSpace(workbookCodeName)) {
                WriteRecord(stream, 0x01ba, BuildCodeNamePayload(workbookCodeName!));
            }

            WriteWorkbookFileSharingRecord(stream, document);

            if (document.DateSystem == ExcelDateSystem.NineteenFour) {
                WriteRecord(stream, 0x0022, BuildUInt16Payload(1));
            }

            WriteWorkbookOptionRecords(stream, document);
            WriteCalculationSettingsRecords(stream, document);
            WriteWorkbookProtectionRecords(stream, document);
            foreach (byte[] windowPayload in BuildWindow1Payloads(document, sheets)) {
                WriteRecord(stream, 0x003d, windowPayload);
            }

            foreach (byte[] fontPayload in fontTable.FontRecords) {
                WriteRecord(stream, 0x0031, fontPayload);
            }

            byte[]? palettePayload = fontTable.PaletteRecord;
            if (palettePayload != null) {
                WriteRecord(stream, 0x0092, palettePayload);
            }

            foreach (byte[] formatPayload in styleTable.FormatRecords) {
                WriteRecord(stream, 0x041e, formatPayload);
            }

            foreach (byte[] cellFormatPayload in styleTable.CellFormatRecords) {
                WriteRecord(stream, 0x00e0, cellFormatPayload);
            }

            foreach (TableStyleRecord tableStyleRecord in CreateTableStyleRecords(document.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet)) {
                WriteRecord(stream, tableStyleRecord.RecordType, tableStyleRecord.Payload);
            }

            var boundSheetPositions = new List<long>(sheets.Length);
            foreach (ExcelSheet sheet in sheets) {
                boundSheetPositions.Add(stream.Position);
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, sheet));
            }

            WriteRecord(stream, 0x013d, BuildSheetTabIdsPayload(document, sheets));
            foreach (LegacyXlsExternSheetTable.SupportingLinkRecord supportingLinkRecord in externSheetTable.SupportingLinkRecords) {
                WriteRecord(stream, supportingLinkRecord.RecordType, supportingLinkRecord.Payload);
            }

            WriteRecord(stream, 0x0017, externSheetTable.Payload);
            if (LegacyXlsAutoFilterWriter.HasWorksheetAutoFilters(sheets) || LegacyXlsDefinedNameWriter.HasWorkbookDefinedNames(document)) {
                foreach (byte[] namePayload in LegacyXlsAutoFilterWriter.CreateDefinedNamePayloads(sheets)) {
                    WriteRecord(stream, 0x0018, namePayload);
                }

                foreach (byte[] namePayload in LegacyXlsDefinedNameWriter.CreateDefinedNamePayloads(document, sheets, formulaNameIndex)) {
                    WriteRecord(stream, 0x0018, namePayload);
                }
            }

            foreach (ExcelSheet sheet in sheets) {
                foreach (byte[] dataConsolidationReferencePayload in LegacyXlsDataConsolidationWriter.CreateReferencePayloads(sheet)) {
                    WriteRecord(stream, 0x0051, dataConsolidationReferencePayload);
                }

                foreach (byte[] dataConsolidationNamePayload in LegacyXlsDataConsolidationWriter.CreateNamePayloads(sheet)) {
                    WriteRecord(stream, 0x0052, dataConsolidationNamePayload);
                }
            }

            WriteRecord(stream, 0x000a, Array.Empty<byte>());

            for (int i = 0; i < sheets.Length; i++) {
                int sheetOffset = checked((int)stream.Position);
                WriteWorksheet(stream, document.WorkbookPartRoot, sheets[i], i, styleTable, document.DateSystem, formulaNameIndex, fontTable);

                long patchOffset = boundSheetPositions[i] + 4;
                long currentPosition = stream.Position;
                stream.Position = patchOffset;
                WriteUInt32(stream, unchecked((uint)sheetOffset));
                stream.Position = currentPosition;
            }

            return stream.ToArray();
        }

        private static void WriteWorksheet(
            Stream stream,
            WorkbookPart workbookPart,
            ExcelSheet sheet,
            int sheetIndex,
            LegacyXlsStyleTable styleTable,
            ExcelDateSystem dateSystem,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            LegacyXlsFontTable fontTable) {
            List<LegacyXlsCell> cells = ExtractCells(sheet, workbookPart, sheetIndex, styleTable, dateSystem, formulaNameIndex, fontTable);
            LegacyXlsWorksheetLayout layout = ExtractWorksheetLayout(sheet, styleTable);
            WriteRecord(stream, 0x0809, WorksheetBof);
            string? worksheetCodeName = sheet.WorksheetPart.Worksheet!.GetFirstChild<SheetProperties>()?.CodeName?.Value;
            if (!string.IsNullOrWhiteSpace(worksheetCodeName)) {
                WriteRecord(stream, 0x01ba, BuildCodeNamePayload(worksheetCodeName!));
            }

            if (layout.DefaultColumnWidth.HasValue) {
                WriteRecord(stream, 0x0055, BuildDefaultColumnWidthPayload(layout.DefaultColumnWidth.Value));
            }

            if (layout.DefaultRowHeight.HasValue) {
                WriteRecord(stream, 0x0225, BuildDefaultRowHeightPayload(layout.DefaultRowHeight.Value, layout.DefaultRowsHidden));
            }

            foreach (LegacyXlsColumnLayout column in layout.Columns) {
                WriteRecord(stream, 0x007d, BuildColInfoPayload(column));
            }

            WriteRecord(stream, 0x0200, BuildDimensionsPayload(cells, layout));

            foreach (LegacyXlsRowLayout row in layout.Rows) {
                WriteRecord(stream, 0x0208, BuildRowPayload(row));
            }

            foreach (LegacyXlsWorksheetView view in layout.Views) {
                WriteRecord(stream, 0x023e, BuildWindow2Payload(view));
            }

            if (TryCreateSheetExtensionPayload(sheet, fontTable, out byte[]? sheetExtensionPayload)) {
                WriteRecord(stream, 0x0862, sheetExtensionPayload!);
            }

            if (layout.View.ZoomScale.HasValue) {
                WriteRecord(stream, 0x00a0, BuildZoomScalePayload(layout.View.ZoomScale.Value));
            }

            if (layout.View.PageLayoutView) {
                WriteRecord(stream, 0x088b, BuildPageLayoutViewPayload(layout.View.ZoomScale));
            }

            if (layout.View.FrozenRowCount > 0 || layout.View.FrozenColumnCount > 0) {
                WriteRecord(stream, 0x0041, BuildPanePayload(layout.View.FrozenColumnCount, layout.View.FrozenRowCount));
            } else if (layout.View.SplitPane.HasValue) {
                WriteRecord(stream, 0x0041, BuildPanePayload(layout.View.SplitPane.Value));
            }

            foreach (LegacyXlsSelection selection in layout.View.Selections) {
                WriteRecord(stream, 0x001d, BuildSelectionPayload(selection));
            }

            byte[]? sortPayload = LegacyXlsSortWriter.CreateSortPayload(sheet);
            if (sortPayload != null) {
                WriteRecord(stream, 0x0090, sortPayload);
            }

            WritePageSetupRecords(stream, layout.PageSetup);
            WriteWorksheetProtectionRecords(stream, layout.Protection);
            if (LegacyXlsWorksheetProtectionFeatureWriter.TryCreatePayload(sheet.WorksheetPart.Worksheet?.Elements<SheetProtection>().FirstOrDefault(), out byte[]? protectionFeaturePayload)) {
                WriteRecord(stream, 0x0867, protectionFeaturePayload!);
            }

            foreach (byte[] protectedRangePayload in LegacyXlsProtectedRangeWriter.CreateProtectedRangePayloads(sheet)) {
                WriteRecord(stream, 0x0868, protectedRangePayload);
            }

            if (LegacyXlsIgnoredErrorWriter.TryCreateHeaderPayload(sheet, out byte[]? ignoredErrorHeaderPayload)) {
                WriteRecord(stream, 0x0867, ignoredErrorHeaderPayload!);
                foreach (byte[] ignoredErrorPayload in LegacyXlsIgnoredErrorWriter.CreateIgnoredErrorPayloads(sheet)) {
                    WriteRecord(stream, 0x0868, ignoredErrorPayload);
                }
            }

            foreach (byte[] cellWatchPayload in LegacyXlsCellWatchWriter.CreateCellWatchPayloads(sheet)) {
                WriteRecord(stream, 0x086c, cellWatchPayload);
            }

            if (LegacyXlsDataConsolidationWriter.TryCreatePayload(sheet, out byte[]? dataConsolidationPayload) && dataConsolidationPayload != null) {
                WriteRecord(stream, 0x0050, dataConsolidationPayload);
            }

            if (LegacyXlsScenarioWriter.TryCreateScenarioManagerPayload(sheet, out byte[]? scenarioManagerPayload)) {
                WriteRecord(stream, 0x00ae, scenarioManagerPayload!);
                foreach (byte[] scenarioPayload in LegacyXlsScenarioWriter.CreateScenarioPayloads(sheet)) {
                    WriteRecord(stream, 0x00af, scenarioPayload);
                }
            }

            WriteWorksheetCalculationRecords(stream, sheet);
            WriteWorksheetPhoneticInfoRecords(stream, sheet);

            foreach (LegacyXlsCell cell in cells) {
                switch (cell.Kind) {
                    case LegacyXlsCellKind.Blank:
                        WriteRecord(stream, 0x0201, BuildBlankPayload(cell.Row, cell.Column, cell.StyleIndex));
                        break;
                    case LegacyXlsCellKind.Number:
                        WriteRecord(stream, 0x0203, BuildNumberPayload(cell.Row, cell.Column, cell.NumberValue, cell.StyleIndex));
                        break;
                    case LegacyXlsCellKind.FormulaNumber:
                        WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(cell.Row, cell.Column, cell.NumberValue, cell.StyleIndex, cell.FormulaTokens, cell.FormulaExtraData));
                        WriteFormulaFollowUpRecords(stream, cell);
                        break;
                    case LegacyXlsCellKind.FormulaBoolean:
                        WriteRecord(stream, 0x0006, BuildFormulaBooleanPayload(cell.Row, cell.Column, cell.BooleanValue, cell.StyleIndex, cell.FormulaTokens, cell.FormulaExtraData));
                        WriteFormulaFollowUpRecords(stream, cell);
                        break;
                    case LegacyXlsCellKind.FormulaText:
                        WriteRecord(stream, 0x0006, BuildFormulaStringPayload(cell.Row, cell.Column, cell.StyleIndex, cell.FormulaTokens, cell.FormulaExtraData));
                        WriteRecord(stream, 0x0207, BuildStringPayload(cell.TextValue ?? string.Empty));
                        WriteFormulaFollowUpRecords(stream, cell);
                        break;
                    case LegacyXlsCellKind.FormulaError:
                        WriteRecord(stream, 0x0006, BuildFormulaErrorPayload(cell.Row, cell.Column, cell.ErrorValue, cell.StyleIndex, cell.FormulaTokens, cell.FormulaExtraData));
                        WriteFormulaFollowUpRecords(stream, cell);
                        break;
                    case LegacyXlsCellKind.Boolean:
                        WriteRecord(stream, 0x0205, BuildBoolErrPayload(cell.Row, cell.Column, cell.BooleanValue, cell.StyleIndex));
                        break;
                    case LegacyXlsCellKind.Error:
                        WriteRecord(stream, 0x0205, BuildErrorPayload(cell.Row, cell.Column, cell.ErrorValue, cell.StyleIndex));
                        break;
                    case LegacyXlsCellKind.Text:
                        WriteRecord(stream, 0x0204, BuildLabelPayload(cell.Row, cell.Column, cell.TextValue ?? string.Empty, cell.StyleIndex, cell.TextFormattingRuns));
                        break;
                }
            }

            if (LegacyXlsAutoFilterWriter.TryCreateAutoFilterInfoPayload(sheet, out byte[]? autoFilterInfoPayload)) {
                WriteRecord(stream, 0x009d, autoFilterInfoPayload!);
                WriteRecord(stream, 0x009b, Array.Empty<byte>());
                foreach (byte[] criteriaPayload in LegacyXlsAutoFilterWriter.CreateCriteriaPayloads(sheet, dateSystem)) {
                    WriteRecord(stream, 0x009e, criteriaPayload);
                }
            }

            if (LegacyXlsDataValidationWriter.TryCreateCollectionPayload(sheet, out byte[]? dataValidationCollectionPayload)) {
                WriteRecord(stream, 0x01b2, dataValidationCollectionPayload!);
                foreach (byte[] dataValidationPayload in LegacyXlsDataValidationWriter.CreateValidationPayloads(sheet, sheetIndex, formulaNameIndex)) {
                    WriteRecord(stream, 0x01be, dataValidationPayload);
                }
            }

            foreach (LegacyXlsConditionalFormattingWriter.ConditionalFormattingBlock conditionalFormattingBlock in LegacyXlsConditionalFormattingWriter.CreateBlocks(sheet, sheetIndex, formulaNameIndex)) {
                WriteRecord(stream, 0x01b0, conditionalFormattingBlock.HeaderPayload);
                foreach (byte[] rulePayload in conditionalFormattingBlock.RulePayloads) {
                    WriteRecord(stream, 0x01b1, rulePayload);
                }

                foreach (byte[] extensionPayload in conditionalFormattingBlock.ExtensionPayloads) {
                    WriteRecord(stream, 0x087b, extensionPayload);
                }
            }

            foreach (LegacyXlsCommentWriter.CommentRecordSet commentRecordSet in LegacyXlsCommentWriter.CreateCommentRecordSets(sheet, fontTable)) {
                WriteRecord(stream, 0x00ec, commentRecordSet.DrawingPayload);
                WriteRecord(stream, 0x005d, commentRecordSet.ObjectPayload);
                WriteRecord(stream, 0x01b6, commentRecordSet.TextObjectPayload);
                WriteRecord(stream, 0x003c, commentRecordSet.TextPayload);
                WriteRecord(stream, 0x003c, commentRecordSet.FormattingPayload);
                WriteRecord(stream, 0x001c, commentRecordSet.NotePayload);
            }

            foreach (LegacyXlsHyperlinkWriter.LegacyXlsHyperlinkRecord hyperlinkRecord in LegacyXlsHyperlinkWriter.CreateHyperlinkRecords(sheet)) {
                WriteRecord(stream, hyperlinkRecord.RecordType, hyperlinkRecord.Payload);
            }

            foreach (byte[] mergeChunk in BuildMergeCellsPayloads(layout.MergedRanges)) {
                WriteRecord(stream, 0x00e5, mergeChunk);
            }

            WriteRecord(stream, 0x000a, Array.Empty<byte>());
        }

        private static List<LegacyXlsCell> ExtractCells(
            ExcelSheet sheet,
            WorkbookPart workbookPart,
            int sheetIndex,
            LegacyXlsStyleTable styleTable,
            ExcelDateSystem dateSystem,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            LegacyXlsFontTable fontTable) {
            SheetData? sheetData = sheet.WorksheetPart.Worksheet?.GetFirstChild<SheetData>();
            var cells = new List<LegacyXlsCell>();
            if (sheetData == null) {
                return cells;
            }

            LegacyXlsSharedFormulaTable sharedFormulaTable = LegacyXlsSharedFormulaTable.Create(sheet, sheetIndex, formulaNameIndex);
            LegacyXlsArrayFormulaTable arrayFormulaTable = LegacyXlsArrayFormulaTable.Create(sheet, sheetIndex, formulaNameIndex);
            foreach (Row row in sheetData.Elements<Row>()) {
                uint rowIndex = row.RowIndex?.Value ?? 0U;
                int sequentialColumn = 1;
                foreach (Cell cell in row.Elements<Cell>()) {
                    uint effectiveRow = rowIndex;
                    int effectiveColumn = sequentialColumn;
                    if (!string.IsNullOrEmpty(cell.CellReference?.Value)) {
                        ParseCellReference(cell.CellReference!.Value!, out effectiveRow, out effectiveColumn);
                    }

                    if (effectiveRow == 0 || effectiveColumn <= 0) {
                        sequentialColumn++;
                        continue;
                    }

                    if (effectiveRow > 65536U || effectiveColumn > 256) {
                        throw new NotSupportedException("Native XLS saving supports the BIFF8 worksheet limit of 65,536 rows and 256 columns.");
                    }

                    LegacyXlsCell? legacyCell = ConvertCell(sheet, workbookPart, sheetIndex, cell, checked((ushort)(effectiveRow - 1U)), checked((ushort)(effectiveColumn - 1)), styleTable, dateSystem, formulaNameIndex, sharedFormulaTable, arrayFormulaTable, fontTable);
                    if (legacyCell.HasValue) {
                        cells.Add(legacyCell.Value);
                    }

                    sequentialColumn = effectiveColumn + 1;
                }
            }

            cells.Sort(static (left, right) => {
                int rowComparison = left.Row.CompareTo(right.Row);
                return rowComparison != 0 ? rowComparison : left.Column.CompareTo(right.Column);
            });
            return cells;
        }

        private static LegacyXlsWorksheetLayout ExtractWorksheetLayout(ExcelSheet sheet, LegacyXlsStyleTable styleTable) {
            var columns = new List<LegacyXlsColumnLayout>();
            foreach (ExcelColumnSnapshot column in sheet.GetColumnDefinitions()) {
                if (column.StartIndex <= 0 || column.EndIndex < column.StartIndex) {
                    continue;
                }

                if (column.EndIndex > 256) {
                    throw new NotSupportedException("Native XLS saving supports the BIFF8 worksheet limit of 256 columns.");
                }

                ushort styleIndex = styleTable.GetBiffStyleIndex(column.StyleIndex);
                columns.Add(new LegacyXlsColumnLayout(
                    checked((ushort)(column.StartIndex - 1)),
                    checked((ushort)(column.EndIndex - 1)),
                    column.Width,
                    column.Hidden,
                    styleIndex,
                    column.OutlineLevel ?? 0,
                    column.Collapsed));
            }

            var rows = new List<LegacyXlsRowLayout>();
            foreach (ExcelRowSnapshot row in sheet.GetRowDefinitions()) {
                if (row.Index <= 0) {
                    continue;
                }

                if (row.Index > 65536) {
                    throw new NotSupportedException("Native XLS saving supports the BIFF8 worksheet limit of 65,536 rows.");
                }

                ushort styleIndex = styleTable.GetBiffStyleIndex(row.StyleIndex);
                rows.Add(new LegacyXlsRowLayout(
                    checked((ushort)(row.Index - 1)),
                    row.Height,
                    row.Hidden,
                    row.CustomHeight,
                    row.CustomFormat,
                    styleIndex,
                    row.OutlineLevel ?? 0,
                    row.Collapsed));
            }

            var mergedRanges = new List<LegacyXlsMergedRange>();
            foreach (ExcelMergedRangeSnapshot mergedRange in sheet.GetMergedRanges()) {
                if (mergedRange.StartRow <= 0 || mergedRange.StartColumn <= 0 || mergedRange.EndRow < mergedRange.StartRow || mergedRange.EndColumn < mergedRange.StartColumn) {
                    continue;
                }

                if (mergedRange.EndRow > 65536 || mergedRange.EndColumn > 256) {
                    throw new NotSupportedException("Native XLS saving supports the BIFF8 worksheet limit of 65,536 rows and 256 columns.");
                }

                mergedRanges.Add(new LegacyXlsMergedRange(
                    checked((ushort)(mergedRange.StartRow - 1)),
                    checked((ushort)(mergedRange.StartColumn - 1)),
                    checked((ushort)(mergedRange.EndRow - 1)),
                    checked((ushort)(mergedRange.EndColumn - 1))));
            }

            ExcelWorksheetViewInfo viewInfo = sheet.GetViewInfo();
            IReadOnlyList<SheetView> sheetViews = sheet.WorksheetPart.Worksheet!
                .GetFirstChild<SheetViews>()?
                .Elements<SheetView>()
                .ToArray() ?? Array.Empty<SheetView>();
            IReadOnlyList<LegacyXlsWorksheetView> views = CreateWorksheetViews(sheet, viewInfo, sheetViews);

            LegacyXlsWorksheetPageSetup pageSetup = ExtractPageSetup(sheet);
            LegacyXlsWorksheetProtection protection = ExtractWorksheetProtection(sheet);

            return new LegacyXlsWorksheetLayout(
                NormalizeDefaultColumnWidth(sheet.DefaultColumnWidth),
                sheet.DefaultRowHeight,
                sheet.DefaultRowsHidden,
                columns,
                rows,
                mergedRanges,
                views,
                pageSetup,
                protection);
        }

        private static IReadOnlyList<LegacyXlsWorksheetView> CreateWorksheetViews(ExcelSheet sheet, ExcelWorksheetViewInfo viewInfo, IReadOnlyList<SheetView> sheetViews) {
            if (sheetViews.Count == 0) {
                return new[] { CreateWorksheetView(sheet, viewInfo, null, primary: true) };
            }

            var views = new List<LegacyXlsWorksheetView>(sheetViews.Count);
            for (int i = 0; i < sheetViews.Count; i++) {
                views.Add(CreateWorksheetView(sheet, viewInfo, sheetViews[i], primary: i == 0));
            }

            return views;
        }

        private static LegacyXlsWorksheetView CreateWorksheetView(ExcelSheet sheet, ExcelWorksheetViewInfo viewInfo, SheetView? sheetView, bool primary) {
            int frozenRowCount = primary ? viewInfo.FrozenRowCount : 0;
            int frozenColumnCount = primary ? viewInfo.FrozenColumnCount : 0;
            LegacyXlsSplitPaneView? splitPane = primary ? ExtractSplitPane(sheetView) : null;
            bool pageBreakPreview = sheetView?.View?.Value == SheetViewValues.PageBreakPreview;
            uint? zoomScale = sheetView?.ZoomScale?.Value ?? (primary ? viewInfo.ZoomScale : null);
            uint? zoomScaleNormal = sheetView?.ZoomScaleNormal?.Value
                ?? (!pageBreakPreview ? zoomScale : null)
                ?? (primary ? viewInfo.ZoomScaleNormal : null);
            return new LegacyXlsWorksheetView(
                frozenRowCount,
                frozenColumnCount,
                sheetView?.ShowFormulas?.Value == true,
                sheetView?.ShowGridLines?.Value ?? (primary ? viewInfo.ShowGridlines : true),
                sheetView?.ShowRowColHeaders?.Value ?? (primary ? sheet.RowColumnHeadingsVisible : true),
                sheetView?.ShowZeros?.Value ?? (primary ? sheet.ZeroValuesVisible : true),
                sheetView?.RightToLeft?.Value ?? (primary && viewInfo.RightToLeft),
                sheetView?.DefaultGridColor?.Value ?? true,
                GetWindow2GridLineColorIndex(sheetView?.DefaultGridColor?.Value ?? true, sheetView?.ColorId?.Value),
                sheetView?.ShowOutlineSymbols?.Value ?? true,
                sheetView?.TabSelected?.Value == true,
                pageBreakPreview,
                sheetView?.View?.Value == SheetViewValues.PageLayout,
                IsFrozenWithoutSplit(sheetView),
                splitPane,
                zoomScale,
                zoomScaleNormal,
                GetWindow2TopLeftCell(sheetView?.TopLeftCell?.Value ?? (primary ? viewInfo.TopLeftCell : null)),
                ExtractSelections(sheetView, frozenRowCount, frozenColumnCount, splitPane));
        }

        private static LegacyXlsWindowTopLeftCell GetWindow2TopLeftCell(string? topLeftCell) {
            if (string.IsNullOrWhiteSpace(topLeftCell)) {
                return new LegacyXlsWindowTopLeftCell(0, 0);
            }

            if (!A1.TryParseCellReferenceFast(topLeftCell, out int row, out int column)) {
                throw new NotSupportedException($"Native XLS saving requires a valid worksheet top-left cell reference; this worksheet uses '{topLeftCell}'.");
            }

            if (row < 1 || row > 65536 || column < 1 || column > 256) {
                throw new NotSupportedException($"Native XLS saving supports top-left visible cells within the BIFF8 worksheet limit A1:IV65536; this worksheet uses '{topLeftCell}'.");
            }

            return new LegacyXlsWindowTopLeftCell(
                checked((ushort)(row - 1)),
                checked((ushort)(column - 1)));
        }

        private static LegacyXlsSplitPaneView? ExtractSplitPane(SheetView? sheetView) {
            Pane? pane = sheetView?.GetFirstChild<Pane>();
            if (pane == null) {
                return null;
            }

            PaneStateValues? state = pane.State?.Value;
            if (state == PaneStateValues.Frozen || state == PaneStateValues.FrozenSplit) {
                return null;
            }

            ushort horizontalSplit = GetSplitPaneCoordinate(pane.HorizontalSplit?.Value, "horizontal split");
            ushort verticalSplit = GetSplitPaneCoordinate(pane.VerticalSplit?.Value, "vertical split");
            if (horizontalSplit == 0 && verticalSplit == 0) {
                return null;
            }

            LegacyXlsWindowTopLeftCell topLeftCell = GetWindow2TopLeftCell(pane.TopLeftCell?.Value);
            return new LegacyXlsSplitPaneView(
                horizontalSplit,
                verticalSplit,
                topLeftCell.Row,
                topLeftCell.Column,
                ToLegacyPane(pane.ActivePane?.Value));
        }

        private static ushort GetSplitPaneCoordinate(double? value, string name) {
            if (!value.HasValue || Math.Abs(value.Value) <= double.Epsilon) {
                return 0;
            }

            if (double.IsNaN(value.Value)
                || double.IsInfinity(value.Value)
                || value.Value < 0D
                || value.Value > ushort.MaxValue
                || Math.Abs(value.Value - Math.Round(value.Value)) > double.Epsilon) {
                throw new NotSupportedException($"Native XLS saving supports integral {name} pane coordinates from 0 through 65,535; this worksheet uses {value.Value}.");
            }

            return checked((ushort)Math.Round(value.Value));
        }

        private static ushort GetWindow2GridLineColorIndex(bool defaultGridColor, uint? colorId) {
            if (defaultGridColor) {
                return 64;
            }

            if (!colorId.HasValue || colorId.Value >= 64U) {
                throw new NotSupportedException("Native XLS saving requires SheetView.ColorId from 0 through 63 when DefaultGridColor is false.");
            }

            return checked((ushort)colorId.Value);
        }

        private static bool IsFrozenWithoutSplit(SheetView? sheetView) {
            Pane? pane = sheetView?.GetFirstChild<Pane>();
            return pane?.State?.Value == PaneStateValues.FrozenSplit;
        }

        private static LegacyXlsCell? ConvertCell(
            ExcelSheet sheet,
            WorkbookPart workbookPart,
            int sheetIndex,
            Cell cell,
            ushort row,
            ushort column,
            LegacyXlsStyleTable styleTable,
            ExcelDateSystem dateSystem,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            LegacyXlsSharedFormulaTable sharedFormulaTable,
            LegacyXlsArrayFormulaTable arrayFormulaTable,
            LegacyXlsFontTable fontTable) {
            DocumentFormat.OpenXml.Spreadsheet.CellValues? dataType = cell.DataType?.Value;
            string rawValue = cell.CellValue?.InnerText ?? string.Empty;
            ushort styleIndex = styleTable.GetBiffStyleIndex(cell.StyleIndex?.Value);

            if (cell.CellFormula != null) {
                return ConvertFormulaCell(sheet, sheetIndex, cell, row, column, styleIndex, dateSystem, formulaNameIndex, sharedFormulaTable, arrayFormulaTable);
            }

            if (arrayFormulaTable.TryGetDefinition(row, column, out LegacyXlsArrayFormulaDefinition arrayDefinition)
                && (arrayDefinition.AnchorRow != row || arrayDefinition.AnchorColumn != column)) {
                return ConvertCachedFormulaCell(sheet, cell, row, column, styleIndex, dateSystem, BuildFormulaReferenceTokens(arrayDefinition.AnchorRow, arrayDefinition.AnchorColumn), Array.Empty<byte>(), null, null, ToA1Address(row, column), isArrayFormula: true);
            }

            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean) {
                return LegacyXlsCell.Boolean(row, column, styleIndex, rawValue == "1" || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase));
            }

            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error) {
                if (LegacyXlsErrorValue.TryGetCode(rawValue, out byte errorCode)) {
                    return LegacyXlsCell.Error(row, column, styleIndex, errorCode);
                }

                throw new NotSupportedException($"Native XLS saving does not support error value '{rawValue}' at {ToA1Address(row, column)}.");
            }

            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Date) {
                if (TryParseOpenXmlDate(rawValue, out DateTime date)) {
                    return LegacyXlsCell.Number(row, column, styleIndex, ExcelDateSystemConverter.ToSerial(date, dateSystem));
                }

                throw new NotSupportedException($"Native XLS saving does not support date value '{rawValue}' at {ToA1Address(row, column)}.");
            }

            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number || dataType == null) {
                if (string.IsNullOrEmpty(rawValue)) {
                    return HasExplicitCellStyle(cell)
                        ? LegacyXlsCell.Blank(row, column, styleIndex)
                        : null;
                }

                if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                    return LegacyXlsCell.Number(row, column, styleIndex, number);
                }
            }

            if (LegacyXlsRichTextCellWriter.TryGetCellRichTextRuns(
                cell,
                workbookPart.SharedStringTablePart?.SharedStringTable,
                fontTable,
                out string? richText,
                out IReadOnlyList<LegacyXlsTextFormattingRun> richTextRuns,
                out string? richTextReason)) {
                if (richTextRuns.Count > 0 && !string.IsNullOrEmpty(richText)) {
                    return LegacyXlsCell.Text(row, column, styleIndex, richText!, richTextRuns);
                }
            } else {
                throw new NotSupportedException($"Native XLS saving does not yet support {richTextReason}. Save as .xlsx or remove this feature before saving as .xls.");
            }

            string text = sheet.GetCellText(cell);
            if (string.IsNullOrEmpty(text)) {
                return HasExplicitCellStyle(cell)
                    ? LegacyXlsCell.Blank(row, column, styleIndex)
                    : null;
            }

            EnsureSupportedLabelTextLength(text, ToA1Address(row, column));
            return LegacyXlsCell.Text(row, column, styleIndex, text);
        }

        private static bool HasExplicitCellStyle(Cell cell) {
            return cell.StyleIndex != null && cell.StyleIndex.HasValue;
        }

        private static LegacyXlsCell ConvertFormulaCell(
            ExcelSheet sheet,
            int sheetIndex,
            Cell cell,
            ushort row,
            ushort column,
            ushort styleIndex,
            ExcelDateSystem dateSystem,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            LegacyXlsSharedFormulaTable sharedFormulaTable,
            LegacyXlsArrayFormulaTable arrayFormulaTable) {
            DocumentFormat.OpenXml.Spreadsheet.CellValues? dataType = cell.DataType?.Value;
            DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues? formulaType = cell.CellFormula?.FormulaType?.Value;
            string formulaText = cell.CellFormula?.Text ?? string.Empty;
            string address = ToA1Address(row, column);
            byte[]? sharedFormulaPayload = null;
            bool isSharedFormula = formulaType == DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared;
            LegacyXlsSharedFormulaDefinition sharedDefinition = default;
            if (isSharedFormula) {
                if (!sharedFormulaTable.TryGetDefinition(cell.CellFormula?.SharedIndex?.Value, out sharedDefinition)) {
                    throw new NotSupportedException($"Native XLS saving could not find a shared formula definition for {address}. Save as .xlsx or remove this formula before saving as .xls.");
                }

                formulaText = sharedDefinition.FormulaText;
                sharedFormulaPayload = sharedDefinition.AnchorRow == row && sharedDefinition.AnchorColumn == column
                    ? sharedDefinition.Payload
                    : null;
            }

            bool isArrayFormula = formulaType == DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Array;
            LegacyXlsArrayFormulaDefinition arrayDefinition = default;
            if (isArrayFormula && !arrayFormulaTable.TryGetDefinition(row, column, out arrayDefinition)) {
                throw new NotSupportedException($"Native XLS saving could not find an array formula definition for {address}. Save as .xlsx or remove this formula before saving as .xls.");
            }

            if (!LegacyXlsFormulaEncoder.TryEncodeFormulaRecord(formulaText, formulaNameIndex, sheetIndex, out byte[] tokens, out byte[] formulaExtraData, out string? reason)) {
                throw new NotSupportedException($"Native XLS saving does not yet support formula '{formulaText}' at {address}: {reason} Save as .xlsx or remove this formula before saving as .xls.");
            }

            byte[]? arrayFormulaPayload = null;
            if (isSharedFormula) {
                tokens = BuildFormulaReferenceTokens(sharedDefinition.AnchorRow, sharedDefinition.AnchorColumn);
                formulaExtraData = Array.Empty<byte>();
            }

            if (isArrayFormula) {
                arrayFormulaPayload = arrayDefinition.Payload;
                tokens = BuildFormulaReferenceTokens(arrayDefinition.AnchorRow, arrayDefinition.AnchorColumn);
                formulaExtraData = Array.Empty<byte>();
            }

            EnsureSupportedFormulaRecordPayloadLength(tokens, formulaExtraData, address);
            return ConvertCachedFormulaCell(sheet, cell, row, column, styleIndex, dateSystem, tokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload, address, isArrayFormula);
        }

        private static LegacyXlsCell ConvertCachedFormulaCell(
            ExcelSheet sheet,
            Cell cell,
            ushort row,
            ushort column,
            ushort styleIndex,
            ExcelDateSystem dateSystem,
            byte[] tokens,
            byte[] formulaExtraData,
            byte[]? arrayFormulaPayload,
            byte[]? sharedFormulaPayload,
            string address,
            bool isArrayFormula) {
            DocumentFormat.OpenXml.Spreadsheet.CellValues? dataType = cell.DataType?.Value;
            string rawValue = cell.CellValue?.InnerText ?? string.Empty;
            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean) {
                if (string.IsNullOrWhiteSpace(rawValue)) {
                    throw new NotSupportedException($"Native XLS saving requires a cached result for formula cell {address}. Enable formula evaluation before saving or save as .xlsx.");
                }

                bool value = rawValue == "1" || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase);
                return LegacyXlsCell.FormulaBoolean(row, column, styleIndex, value, tokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload);
            }

            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                || dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                || dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                string text = sheet.GetCellText(cell);
                EnsureSupportedFormulaTextLength(text, address);
                return LegacyXlsCell.FormulaText(
                    row,
                    column,
                    styleIndex,
                    text,
                    tokens,
                    formulaExtraData,
                    arrayFormulaPayload: arrayFormulaPayload,
                    sharedFormulaPayload: sharedFormulaPayload);
            }

            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error) {
                if (string.IsNullOrWhiteSpace(rawValue)) {
                    throw new NotSupportedException($"Native XLS saving requires a cached result for formula cell {address}. Enable formula evaluation before saving or save as .xlsx.");
                }

                if (LegacyXlsErrorValue.TryGetCode(rawValue, out byte errorCode)) {
                    return LegacyXlsCell.FormulaError(row, column, styleIndex, errorCode, tokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload);
                }

                throw new NotSupportedException($"Native XLS saving does not support cached error result '{rawValue}' for formula cell {address}.");
            }

            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Date) {
                if (string.IsNullOrWhiteSpace(rawValue)) {
                    throw new NotSupportedException($"Native XLS saving requires a cached result for formula cell {address}. Enable formula evaluation before saving or save as .xlsx.");
                }

                if (TryParseOpenXmlDate(rawValue, out DateTime date)) {
                    return LegacyXlsCell.FormulaNumber(row, column, styleIndex, ExcelDateSystemConverter.ToSerial(date, dateSystem), tokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload);
                }

                throw new NotSupportedException($"Native XLS saving does not support cached date result '{rawValue}' for formula cell {address}.");
            }

            if ((dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number || dataType == null)
                && !string.IsNullOrWhiteSpace(rawValue)
                && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                return LegacyXlsCell.FormulaNumber(row, column, styleIndex, number, tokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload);
            }

            throw new NotSupportedException($"Native XLS saving requires a numeric, Boolean, text, error, or date cached result for formula cell {address}. Enable formula evaluation before saving or save as .xlsx.");
        }

        private static void EnsureSupportedLabelTextLength(string text, string address) {
            if (text.Length > ushort.MaxValue || 9L + GetEncodedStringByteCount(text) > ushort.MaxValue) {
                throw new NotSupportedException($"Native XLS saving does not yet support cell text lengths outside BIFF8 limits at {address}. Save as .xlsx or shorten this cell before saving as .xls.");
            }
        }

        private static void EnsureSupportedFormulaTextLength(string text, string address) {
            if (text.Length > ushort.MaxValue || 3L + GetEncodedStringByteCount(text) > ushort.MaxValue) {
                throw new NotSupportedException($"Native XLS saving does not yet support cached formula text lengths outside BIFF8 limits at {address}. Save as .xlsx or shorten this cached result before saving as .xls.");
            }
        }

        private static void EnsureSupportedFormulaRecordPayloadLength(byte[] formulaTokens, byte[] formulaExtraData, string address) {
            EnsureSupportedFormulaPayloadLength(formulaTokens, formulaExtraData.Length, 22, "formula", address);
        }

        private static void EnsureSupportedArrayFormulaPayloadLength(byte[] formulaTokens, string address) {
            EnsureSupportedFormulaPayloadLength(formulaTokens, extraPayloadLength: 0, fixedPayloadLength: 14, "array formula", address);
        }

        private static void EnsureSupportedSharedFormulaPayloadLength(byte[] formulaTokens, string address) {
            EnsureSupportedFormulaPayloadLength(formulaTokens, extraPayloadLength: 0, fixedPayloadLength: 10, "shared formula", address);
        }

        private static void EnsureSupportedFormulaPayloadLength(byte[] formulaTokens, int extraPayloadLength, int fixedPayloadLength, string context, string address) {
            if (formulaTokens.Length > ushort.MaxValue || fixedPayloadLength + (long)formulaTokens.Length + extraPayloadLength > ushort.MaxValue) {
                throw new NotSupportedException($"Native XLS saving does not yet support {context} token payload lengths outside BIFF8 limits at {address}. Save as .xlsx or shorten this formula before saving as .xls.");
            }
        }

        private static bool TryParseOpenXmlDate(string text, out DateTime value) {
            return DateTime.TryParse(
                text,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind,
                out value);
        }

        private static string ToA1Address(ushort zeroBasedRow, ushort zeroBasedColumn) {
            int column = zeroBasedColumn + 1;
            var letters = new StringBuilder();
            while (column > 0) {
                column--;
                letters.Insert(0, (char)('A' + (column % 26)));
                column /= 26;
            }

            return letters + (zeroBasedRow + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static bool TryParseFormulaRange(string? reference, out ushort firstRow, out ushort firstColumn, out ushort lastRow, out ushort lastColumn) {
            firstRow = firstColumn = lastRow = lastColumn = 0;
            if (string.IsNullOrWhiteSpace(reference)) {
                return false;
            }

            string normalizedReference = reference!.Trim().Replace("$", string.Empty);
            int r1;
            int c1;
            int r2;
            int c2;
            if (normalizedReference.IndexOf(':') < 0) {
                (r1, c1) = A1.ParseCellRef(normalizedReference);
                r2 = r1;
                c2 = c1;
            } else if (!A1.TryParseRange(normalizedReference, out r1, out c1, out r2, out c2)) {
                return false;
            }

            if (r1 <= 0 || c1 <= 0 || r2 > 65536 || c2 > 256) {
                return false;
            }

            firstRow = checked((ushort)(r1 - 1));
            firstColumn = checked((ushort)(c1 - 1));
            lastRow = checked((ushort)(r2 - 1));
            lastColumn = checked((ushort)(c2 - 1));
            return true;
        }

        private static void ParseCellReference(string reference, out uint row, out int column) {
            row = 0;
            column = 0;
            for (int i = 0; i < reference.Length; i++) {
                char ch = reference[i];
                if (ch >= 'A' && ch <= 'Z') {
                    column = checked(column * 26 + (ch - 'A' + 1));
                    continue;
                }

                if (ch >= 'a' && ch <= 'z') {
                    column = checked(column * 26 + (ch - 'a' + 1));
                    continue;
                }

                if (ch >= '0' && ch <= '9') {
                    row = checked(row * 10U + (uint)(ch - '0'));
                }
            }
        }

        private static void EnsureSingleCellArrayFormulaRange(string? reference, ushort row, ushort column, string address) {
            int expectedRow = row + 1;
            int expectedColumn = column + 1;
            if (string.IsNullOrWhiteSpace(reference)) {
                return;
            }

            string normalizedReference = reference!.Trim().Replace("$", string.Empty);
            int separator = normalizedReference.IndexOf(':');
            int firstRow;
            int firstColumn;
            int lastRow;
            int lastColumn;
            if (separator < 0) {
                (firstRow, firstColumn) = A1.ParseCellRef(normalizedReference);
                lastRow = firstRow;
                lastColumn = firstColumn;
            } else if (!A1.TryParseRange(normalizedReference, out firstRow, out firstColumn, out lastRow, out lastColumn)) {
                throw new NotSupportedException($"Native XLS saving does not support array formula range '{reference}' at {address}. Save as .xlsx or remove this formula before saving as .xls.");
            }

            if (firstRow == 0 || firstColumn == 0) {
                throw new NotSupportedException($"Native XLS saving does not support array formula range '{reference}' at {address}. Save as .xlsx or remove this formula before saving as .xls.");
            }

            if (firstRow != lastRow || firstColumn != lastColumn) {
                throw new NotSupportedException($"Native XLS saving supports single-cell array formulas only; multi-cell array formulas such as '{reference}' at {address} are not yet supported. Save as .xlsx or remove this formula before saving as .xls.");
            }

            if (firstRow != expectedRow || firstColumn != expectedColumn) {
                throw new NotSupportedException($"Native XLS saving requires array formula range '{reference}' to match formula cell {address}. Save as .xlsx or remove this formula before saving as .xls.");
            }
        }

        private static byte[] BuildBoundSheetPayload(int streamOffset, ExcelSheet sheet) {
            string sheetName = string.IsNullOrEmpty(sheet.Name) ? "Sheet1" : sheet.Name;
            if (sheetName.Length > 31) {
                throw new NotSupportedException("Native XLS saving supports worksheet names up to 31 characters.");
            }

            byte[] nameBytes = EncodeShortUnicodeString(sheetName, out byte flags);
            byte[] payload = new byte[checked(8 + nameBytes.Length)];
            WriteUInt32(payload, 0, unchecked((uint)streamOffset));
            payload[4] = sheet.VeryHidden ? (byte)2 : sheet.Hidden ? (byte)1 : (byte)0;
            payload[5] = 0;
            payload[6] = checked((byte)sheetName.Length);
            payload[7] = flags;
            Buffer.BlockCopy(nameBytes, 0, payload, 8, nameBytes.Length);
            return payload;
        }

        private static byte[] BuildDimensionsPayload(IReadOnlyList<LegacyXlsCell> cells, LegacyXlsWorksheetLayout layout) {
            uint? firstRow = null;
            uint? lastRow = null;
            ushort? firstColumn = null;
            ushort? lastColumn = null;

            foreach (LegacyXlsCell cell in cells) {
                IncludeDimension(ref firstRow, ref lastRow, ref firstColumn, ref lastColumn, cell.Row, cell.Column);
            }

            foreach (LegacyXlsRowLayout row in layout.Rows) {
                IncludeDimension(ref firstRow, ref lastRow, ref firstColumn, ref lastColumn, row.Row, 0);
            }

            foreach (LegacyXlsColumnLayout column in layout.Columns) {
                IncludeDimension(ref firstRow, ref lastRow, ref firstColumn, ref lastColumn, 0, column.FirstColumn);
                IncludeDimension(ref firstRow, ref lastRow, ref firstColumn, ref lastColumn, 0, column.LastColumn);
            }

            foreach (LegacyXlsMergedRange mergedRange in layout.MergedRanges) {
                IncludeDimension(ref firstRow, ref lastRow, ref firstColumn, ref lastColumn, mergedRange.FirstRow, mergedRange.FirstColumn);
                IncludeDimension(ref firstRow, ref lastRow, ref firstColumn, ref lastColumn, mergedRange.LastRow, mergedRange.LastColumn);
            }

            if (!firstRow.HasValue || !lastRow.HasValue || !firstColumn.HasValue || !lastColumn.HasValue) {
                return BuildDimensionsPayload(0, 0, 0, 0);
            }

            return BuildDimensionsPayload(firstRow.Value, lastRow.Value + 1U, firstColumn.Value, checked((ushort)(lastColumn.Value + 1)));
        }

        private static void IncludeDimension(ref uint? firstRow, ref uint? lastRow, ref ushort? firstColumn, ref ushort? lastColumn, ushort row, ushort column) {
            firstRow = firstRow.HasValue ? Math.Min(firstRow.Value, row) : row;
            lastRow = lastRow.HasValue ? Math.Max(lastRow.Value, row) : row;
            firstColumn = firstColumn.HasValue ? (ushort)Math.Min(firstColumn.Value, column) : column;
            lastColumn = lastColumn.HasValue ? (ushort)Math.Max(lastColumn.Value, column) : column;
        }

        private static byte[] BuildDimensionsPayload(uint firstRow, uint rowAfterLast, ushort firstColumn, ushort columnAfterLast) {
            using var stream = new MemoryStream();
            WriteUInt32(stream, firstRow);
            WriteUInt32(stream, rowAfterLast);
            WriteUInt16(stream, firstColumn);
            WriteUInt16(stream, columnAfterLast);
            WriteUInt16(stream, 0);
            return stream.ToArray();
        }

        private static byte[] BuildLabelPayload(ushort row, ushort column, string text, ushort styleIndex, IReadOnlyList<LegacyXlsTextFormattingRun>? formattingRuns = null) {
            IReadOnlyList<LegacyXlsTextFormattingRun> runs = formattingRuns ?? Array.Empty<LegacyXlsTextFormattingRun>();
            byte[] textBytes = EncodeUnicodeString(text, out byte flags);
            if (runs.Count > 0) {
                flags |= 0x08;
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, row);
            WriteUInt16(stream, column);
            WriteUInt16(stream, styleIndex);
            WriteUInt16(stream, checked((ushort)text.Length));
            stream.WriteByte(flags);
            if (runs.Count > 0) {
                WriteUInt16(stream, checked((ushort)runs.Count));
            }

            stream.Write(textBytes, 0, textBytes.Length);
            foreach (LegacyXlsTextFormattingRun run in runs) {
                WriteUInt16(stream, run.StartCharacter);
                WriteUInt16(stream, run.FontIndex);
            }

            return stream.ToArray();
        }

        private static byte[] BuildBlankPayload(ushort row, ushort column, ushort styleIndex) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, row);
            WriteUInt16(stream, column);
            WriteUInt16(stream, styleIndex);
            return stream.ToArray();
        }

        private static byte[] BuildNumberPayload(ushort row, ushort column, double value, ushort styleIndex) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, row);
            WriteUInt16(stream, column);
            WriteUInt16(stream, styleIndex);
            byte[] number = BitConverter.GetBytes(value);
            stream.Write(number, 0, number.Length);
            return stream.ToArray();
        }

        private static byte[] BuildFormulaNumberPayload(ushort row, ushort column, double value, ushort styleIndex, byte[] formulaTokens, byte[] formulaExtraData) {
            byte[] payload = BuildFormulaPayload(row, column, styleIndex, formulaTokens, formulaExtraData);
            byte[] numberBytes = BitConverter.GetBytes(value);
            Buffer.BlockCopy(numberBytes, 0, payload, 6, numberBytes.Length);
            return payload;
        }

        private static byte[] BuildFormulaBooleanPayload(ushort row, ushort column, bool value, ushort styleIndex, byte[] formulaTokens, byte[] formulaExtraData) {
            byte[] payload = BuildFormulaSpecialPayload(row, column, 0x01, value ? (byte)1 : (byte)0, styleIndex, formulaTokens, formulaExtraData);
            return payload;
        }

        private static byte[] BuildFormulaStringPayload(ushort row, ushort column, ushort styleIndex, byte[] formulaTokens, byte[] formulaExtraData) {
            return BuildFormulaSpecialPayload(row, column, valueType: 0x00, value: 0, styleIndex, formulaTokens, formulaExtraData);
        }

        private static byte[] BuildFormulaErrorPayload(ushort row, ushort column, byte errorCode, ushort styleIndex, byte[] formulaTokens, byte[] formulaExtraData) {
            return BuildFormulaSpecialPayload(row, column, valueType: 0x02, value: errorCode, styleIndex, formulaTokens, formulaExtraData);
        }

        private static byte[] BuildFormulaSpecialPayload(ushort row, ushort column, byte valueType, byte value, ushort styleIndex, byte[] formulaTokens, byte[] formulaExtraData) {
            byte[] payload = BuildFormulaPayload(row, column, styleIndex, formulaTokens, formulaExtraData);
            payload[6] = valueType;
            payload[8] = value;
            WriteUInt16(payload, 12, 0xffff);
            return payload;
        }

        private static byte[] BuildFormulaPayload(ushort row, ushort column, ushort styleIndex, byte[] formulaTokens, byte[] formulaExtraData) {
            byte[] payload = new byte[checked(22 + formulaTokens.Length + formulaExtraData.Length)];
            WriteUInt16(payload, 0, row);
            WriteUInt16(payload, 2, column);
            WriteUInt16(payload, 4, styleIndex);
            WriteUInt16(payload, 20, checked((ushort)formulaTokens.Length));
            Buffer.BlockCopy(formulaTokens, 0, payload, 22, formulaTokens.Length);
            Buffer.BlockCopy(formulaExtraData, 0, payload, 22 + formulaTokens.Length, formulaExtraData.Length);
            return payload;
        }

        private static byte[] BuildFormulaReferenceTokens(ushort row, ushort column) {
            byte[] tokens = new byte[5];
            tokens[0] = 0x01;
            WriteUInt16(tokens, 1, row);
            WriteUInt16(tokens, 3, column);
            return tokens;
        }

        private static byte[] BuildArrayFormulaPayload(ushort row, ushort column, byte[] formulaTokens) {
            return BuildArrayFormulaPayload(row, column, row, column, formulaTokens);
        }

        private static byte[] BuildArrayFormulaPayload(ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn, byte[] formulaTokens) {
            byte[] payload = new byte[checked(14 + formulaTokens.Length)];
            WriteUInt16(payload, 0, firstRow);
            WriteUInt16(payload, 2, lastRow);
            payload[4] = checked((byte)firstColumn);
            payload[5] = checked((byte)lastColumn);
            WriteUInt16(payload, 12, checked((ushort)formulaTokens.Length));
            Buffer.BlockCopy(formulaTokens, 0, payload, 14, formulaTokens.Length);
            return payload;
        }

        private static byte[] BuildSharedFormulaPayload(ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn, byte[] formulaTokens) {
            byte[] payload = new byte[checked(10 + formulaTokens.Length)];
            WriteUInt16(payload, 0, firstRow);
            WriteUInt16(payload, 2, lastRow);
            payload[4] = checked((byte)firstColumn);
            payload[5] = checked((byte)lastColumn);
            payload[7] = checked((byte)(lastRow - firstRow + 1));
            WriteUInt16(payload, 8, checked((ushort)formulaTokens.Length));
            Buffer.BlockCopy(formulaTokens, 0, payload, 10, formulaTokens.Length);
            return payload;
        }

        private static bool TryBuildSharedFormulaTokens(byte[] formulaTokens, ushort anchorRow, ushort anchorColumn, out byte[] sharedFormulaTokens) {
            sharedFormulaTokens = new byte[formulaTokens.Length];
            Buffer.BlockCopy(formulaTokens, 0, sharedFormulaTokens, 0, formulaTokens.Length);

            int offset = 0;
            while (offset < sharedFormulaTokens.Length) {
                byte token = sharedFormulaTokens[offset];
                int tokenStart = offset++;
                switch (token) {
                    case >= 0x03 and <= 0x16:
                        break;
                    case 0x17:
                        if (offset + 2 > sharedFormulaTokens.Length) return false;
                        int characterCount = sharedFormulaTokens[offset];
                        byte flags = sharedFormulaTokens[offset + 1];
                        offset += checked(2 + (((flags & 0x01) != 0) ? characterCount * 2 : characterCount));
                        break;
                    case 0x19:
                        offset += 3;
                        break;
                    case 0x1c:
                    case 0x1d:
                        offset += 1;
                        break;
                    case 0x1e:
                        offset += 2;
                        break;
                    case 0x1f:
                        offset += 8;
                        break;
                    case 0x41:
                        offset += 2;
                        break;
                    case 0x42:
                        offset += 3;
                        break;
                    case 0x43:
                        offset += 4;
                        break;
                    case 0x44:
                        if (offset + 4 > sharedFormulaTokens.Length) return false;
                        sharedFormulaTokens[tokenStart] = 0x4c;
                        ConvertSharedFormulaReference(sharedFormulaTokens, offset, offset + 2, anchorRow, anchorColumn);
                        offset += 4;
                        break;
                    case 0x45:
                        if (offset + 8 > sharedFormulaTokens.Length) return false;
                        sharedFormulaTokens[tokenStart] = 0x4d;
                        ConvertSharedFormulaReference(sharedFormulaTokens, offset, offset + 4, anchorRow, anchorColumn);
                        ConvertSharedFormulaReference(sharedFormulaTokens, offset + 2, offset + 6, anchorRow, anchorColumn);
                        offset += 8;
                        break;
                    case 0x5a:
                        offset += 6;
                        break;
                    case 0x5b:
                        offset += 10;
                        break;
                    default:
                        return false;
                }

                if (offset > sharedFormulaTokens.Length) {
                    return false;
                }
            }

            return offset == sharedFormulaTokens.Length;
        }

        private static void ConvertSharedFormulaReference(byte[] tokens, int rowOffset, int columnOffset, ushort anchorRow, ushort anchorColumn) {
            ushort row = ReadUInt16(tokens, rowOffset);
            ushort columnBits = ReadUInt16(tokens, columnOffset);
            if ((columnBits & 0x8000) != 0) {
                WriteUInt16(tokens, rowOffset, unchecked((ushort)(short)(row - anchorRow)));
            }

            if ((columnBits & 0x4000) != 0) {
                short relativeColumnOffset = unchecked((short)((columnBits & 0x3fff) - anchorColumn));
                ushort updatedColumnBits = (ushort)((columnBits & 0xc000) | (((ushort)relativeColumnOffset) & 0x3fff));
                WriteUInt16(tokens, columnOffset, updatedColumnBits);
            }
        }

        private static byte[] BuildStringPayload(string text) {
            byte[] textBytes = EncodeUnicodeString(text, out byte flags);
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)text.Length));
            stream.WriteByte(flags);
            stream.Write(textBytes, 0, textBytes.Length);
            return stream.ToArray();
        }

        private static byte[] BuildBoolErrPayload(ushort row, ushort column, bool value, ushort styleIndex) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, row);
            WriteUInt16(stream, column);
            WriteUInt16(stream, styleIndex);
            stream.WriteByte(value ? (byte)1 : (byte)0);
            stream.WriteByte(0);
            return stream.ToArray();
        }

        private static byte[] BuildErrorPayload(ushort row, ushort column, byte errorCode, ushort styleIndex) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, row);
            WriteUInt16(stream, column);
            WriteUInt16(stream, styleIndex);
            stream.WriteByte(errorCode);
            stream.WriteByte(1);
            return stream.ToArray();
        }

        private static byte[] BuildDefaultColumnWidthPayload(double width) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)Math.Round(width)));
            return stream.ToArray();
        }

        private static byte[] BuildDefaultRowHeightPayload(double height, bool hidden) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, hidden ? (ushort)0x0002 : (ushort)0x0000);
            WriteUInt16(stream, checked((ushort)Math.Round(height * 20d)));
            return stream.ToArray();
        }

        private static byte[] BuildColInfoPayload(LegacyXlsColumnLayout column) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, column.FirstColumn);
            WriteUInt16(stream, column.LastColumn);
            WriteUInt16(stream, checked((ushort)Math.Round((column.Width ?? 8.43d) * 256d)));
            WriteUInt16(stream, column.StyleIndex);
            ushort options = column.Hidden ? (ushort)0x0001 : (ushort)0x0000;
            options |= (ushort)((column.OutlineLevel & 0x07) << 8);
            if (column.Collapsed) {
                options |= 0x1000;
            }

            WriteUInt16(stream, options);
            WriteUInt16(stream, 0);
            return stream.ToArray();
        }

        private static byte[] BuildRowPayload(LegacyXlsRowLayout row) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, row.Row);
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 1);
            WriteUInt16(stream, checked((ushort)Math.Round((row.Height ?? 15d) * 20d)));
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 0);
            ushort options = (ushort)(row.OutlineLevel & 0x07);
            if (row.Collapsed) {
                options |= 0x0010;
            }

            if (row.Hidden) {
                options |= 0x0020;
            }

            if (row.CustomHeight || row.Height.HasValue) {
                options |= 0x0040;
            }

            if (row.CustomFormat) {
                options |= 0x0080;
            }

            WriteUInt16(stream, options);
            WriteUInt16(stream, row.CustomFormat ? (ushort)(row.StyleIndex & 0x0fff) : (ushort)0x0100);
            return stream.ToArray();
        }

        private static IEnumerable<byte[]> BuildMergeCellsPayloads(IReadOnlyList<LegacyXlsMergedRange> ranges) {
            const int MaxRangesPerRecord = 1027;
            for (int offset = 0; offset < ranges.Count; offset += MaxRangesPerRecord) {
                int count = Math.Min(MaxRangesPerRecord, ranges.Count - offset);
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)count));
                for (int i = 0; i < count; i++) {
                    LegacyXlsMergedRange range = ranges[offset + i];
                    WriteUInt16(stream, range.FirstRow);
                    WriteUInt16(stream, range.LastRow);
                    WriteUInt16(stream, range.FirstColumn);
                    WriteUInt16(stream, range.LastColumn);
                }

                yield return stream.ToArray();
            }
        }

        private static LegacyXlsWorksheetPageSetup ExtractPageSetup(ExcelSheet sheet) {
            ExcelSheetPageSetup pageSetup = sheet.GetPageSetup();
            ExcelSheetPrintOptions printOptions = sheet.GetPrintOptions();
            ExcelSheet.HeaderFooterSnapshot headerFooter = sheet.GetHeaderFooter();
            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            SheetProperties? sheetProperties = worksheet.GetFirstChild<SheetProperties>();
            OutlineProperties? outlineProperties = sheetProperties?.GetFirstChild<OutlineProperties>();
            PageSetupProperties? pageSetupProperties = sheetProperties?.GetFirstChild<PageSetupProperties>();
            HeaderFooter? headerFooterElement = worksheet.GetFirstChild<HeaderFooter>();
            PrintOptions? printOptionsElement = worksheet.GetFirstChild<PrintOptions>();
            bool? fitToPage = pageSetupProperties?.FitToPage?.Value;
            if (!fitToPage.HasValue && (pageSetup.FitToWidth.HasValue || pageSetup.FitToHeight.HasValue)) {
                fitToPage = true;
            }

            return new LegacyXlsWorksheetPageSetup(
                pageSetup.Orientation,
                pageSetup.Margins?.Left,
                pageSetup.Margins?.Right,
                pageSetup.Margins?.Top,
                pageSetup.Margins?.Bottom,
                pageSetup.Margins?.Header,
                pageSetup.Margins?.Footer,
                pageSetup.FitToWidth,
                pageSetup.FitToHeight,
                pageSetup.Scale,
                pageSetup.PageOrder,
                fitToPage,
                printOptions.PrintGridLines,
                printOptions.PrintHeadings,
                printOptions.HorizontalCentered,
                printOptions.VerticalCentered,
                printOptionsElement?.GridLinesSet?.Value,
                outlineProperties?.ApplyStyles?.Value,
                outlineProperties?.SummaryBelow?.Value,
                outlineProperties?.SummaryRight?.Value,
                BuildHeaderFooterText(headerFooter.HeaderLeft, headerFooter.HeaderCenter, headerFooter.HeaderRight),
                BuildHeaderFooterText(headerFooter.FooterLeft, headerFooter.FooterCenter, headerFooter.FooterRight),
                BuildHeaderFooterText(headerFooter.FirstHeaderLeft, headerFooter.FirstHeaderCenter, headerFooter.FirstHeaderRight),
                BuildHeaderFooterText(headerFooter.FirstFooterLeft, headerFooter.FirstFooterCenter, headerFooter.FirstFooterRight),
                BuildHeaderFooterText(headerFooter.EvenHeaderLeft, headerFooter.EvenHeaderCenter, headerFooter.EvenHeaderRight),
                BuildHeaderFooterText(headerFooter.EvenFooterLeft, headerFooter.EvenFooterCenter, headerFooter.EvenFooterRight),
                headerFooter.DifferentFirstPage,
                headerFooter.DifferentOddEven,
                headerFooterElement?.ScaleWithDoc?.Value ?? true,
                headerFooterElement?.AlignWithMargins?.Value ?? true,
                sheet.GetManualRowPageBreaks(),
                sheet.GetManualColumnPageBreaks(),
                ExtractPrinterSettingsPayload(sheet.WorksheetPart));
        }

        private static void WritePageSetupRecords(Stream stream, LegacyXlsWorksheetPageSetup pageSetup) {
            if (pageSetup.PrintHeadings.HasValue) {
                WriteRecord(stream, 0x002a, BuildUInt16Payload(pageSetup.PrintHeadings.Value ? (ushort)1 : (ushort)0));
            }

            if (pageSetup.PrintGridLines.HasValue) {
                WriteRecord(stream, 0x002b, BuildUInt16Payload(pageSetup.PrintGridLines.Value ? (ushort)1 : (ushort)0));
            }

            if (pageSetup.GridLinesSet.HasValue) {
                WriteRecord(stream, 0x0082, BuildUInt16Payload(pageSetup.GridLinesSet.Value ? (ushort)1 : (ushort)0));
            }

            if (pageSetup.HorizontalCentered.HasValue) {
                WriteRecord(stream, 0x0083, BuildUInt16Payload(pageSetup.HorizontalCentered.Value ? (ushort)1 : (ushort)0));
            }

            if (pageSetup.VerticalCentered.HasValue) {
                WriteRecord(stream, 0x0084, BuildUInt16Payload(pageSetup.VerticalCentered.Value ? (ushort)1 : (ushort)0));
            }

            if (pageSetup.LeftMargin.HasValue) {
                WriteRecord(stream, 0x0026, BuildDoublePayload(pageSetup.LeftMargin.Value));
            }

            if (pageSetup.RightMargin.HasValue) {
                WriteRecord(stream, 0x0027, BuildDoublePayload(pageSetup.RightMargin.Value));
            }

            if (pageSetup.TopMargin.HasValue) {
                WriteRecord(stream, 0x0028, BuildDoublePayload(pageSetup.TopMargin.Value));
            }

            if (pageSetup.BottomMargin.HasValue) {
                WriteRecord(stream, 0x0029, BuildDoublePayload(pageSetup.BottomMargin.Value));
            }

            if (pageSetup.PrinterSettingsPayload != null) {
                WriteRecord(stream, 0x004d, pageSetup.PrinterSettingsPayload);
            }

            if (pageSetup.HasSetupRecord) {
                WriteRecord(stream, 0x00a1, BuildSetupPayload(pageSetup));
            }

            if (TryBuildWorksheetOptionsPayload(pageSetup, out byte[]? worksheetOptionsPayload)) {
                WriteRecord(stream, 0x0081, worksheetOptionsPayload!);
            }

            if (!string.IsNullOrEmpty(pageSetup.HeaderText)) {
                WriteRecord(stream, 0x0014, BuildUnicodeStringPayload(pageSetup.HeaderText!));
            }

            if (!string.IsNullOrEmpty(pageSetup.FooterText)) {
                WriteRecord(stream, 0x0015, BuildUnicodeStringPayload(pageSetup.FooterText!));
            }

            if (pageSetup.HasHeaderFooterExtensionRecord) {
                WriteRecord(stream, 0x089c, BuildHeaderFooterExtensionPayload(pageSetup));
            }

            if (pageSetup.RowPageBreaks.Count > 0) {
                WriteRecord(stream, 0x001b, BuildHorizontalPageBreaksPayload(pageSetup.RowPageBreaks));
            }

            if (pageSetup.ColumnPageBreaks.Count > 0) {
                WriteRecord(stream, 0x001a, BuildVerticalPageBreaksPayload(pageSetup.ColumnPageBreaks));
            }
        }

        private static bool TryBuildWorksheetOptionsPayload(LegacyXlsWorksheetPageSetup pageSetup, out byte[]? payload) {
            ushort options = 0;
            bool hasExplicitOption = false;

            if (pageSetup.ApplyOutlineStyles.HasValue) {
                hasExplicitOption = true;
                if (pageSetup.ApplyOutlineStyles.Value) {
                    options |= 0x0020;
                }
            }

            if (pageSetup.SummaryRowsBelow.HasValue) {
                hasExplicitOption = true;
                if (pageSetup.SummaryRowsBelow.Value) {
                    options |= 0x0040;
                }
            }

            if (pageSetup.SummaryColumnsRight.HasValue) {
                hasExplicitOption = true;
                if (pageSetup.SummaryColumnsRight.Value) {
                    options |= 0x0080;
                }
            }

            if (pageSetup.FitToPage.HasValue) {
                hasExplicitOption = true;
                if (pageSetup.FitToPage.Value) {
                    options |= 0x0100;
                }
            }

            payload = hasExplicitOption ? BuildUInt16Payload(options) : null;
            return hasExplicitOption;
        }

        private static byte[]? ExtractPrinterSettingsPayload(WorksheetPart worksheetPart) {
            SpreadsheetPrinterSettingsPart? printerSettingsPart = worksheetPart.SpreadsheetPrinterSettingsParts.FirstOrDefault();
            if (printerSettingsPart == null) {
                return null;
            }

            using Stream source = printerSettingsPart.GetStream(FileMode.Open, FileAccess.Read);
            using var printerSettings = new MemoryStream();
            source.CopyTo(printerSettings);
            return BuildPrinterSettingsPayload(printerSettings.ToArray());
        }

        private static void WriteWorkbookProtectionRecords(Stream stream, ExcelDocument document) {
            WorkbookProtection? protection = document.WorkbookRoot.GetFirstChild<WorkbookProtection>();
            if (protection == null) {
                return;
            }

            if (protection.LockStructure?.Value == true) {
                WriteRecord(stream, 0x0012, BuildUInt16Payload(1));
            }

            ushort? passwordHash = ParseLegacyHashOrThrow(protection.WorkbookPassword?.Value, "workbook protection password");
            if (passwordHash.HasValue) {
                WriteRecord(stream, 0x0013, BuildUInt16Payload(passwordHash.Value));
            }

            if (protection.LockWindows?.Value == true) {
                WriteRecord(stream, 0x0019, BuildUInt16Payload(1));
            }

            if (protection.LockRevision?.Value is bool lockRevision) {
                WriteRecord(stream, 0x01af, BuildUInt16Payload(lockRevision ? (ushort)1 : (ushort)0));
            }

            ushort? revisionPasswordHash = ParseLegacyHashOrThrow(protection.RevisionsPassword?.Value, "workbook revision protection password");
            if (revisionPasswordHash.HasValue) {
                WriteRecord(stream, 0x01bc, BuildUInt16Payload(revisionPasswordHash.Value));
            }
        }

        private static void WriteWorkbookFileSharingRecord(Stream stream, ExcelDocument document) {
            FileSharing? fileSharing = document.WorkbookRoot.GetFirstChild<FileSharing>();
            if (fileSharing == null) {
                return;
            }

            WriteRecord(stream, 0x005b, BuildFileSharingPayload(fileSharing));
            ushort? passwordHash = ParseLegacyHashOrThrow(GetFileSharingReservationPassword(fileSharing), "write-reservation password");
            string? userName = fileSharing.UserName?.Value;
            if (!passwordHash.HasValue && !string.IsNullOrWhiteSpace(userName)) {
                WriteRecord(stream, 0x005c, BuildWriteAccessPayload(userName!));
            }
        }

        private static byte[] BuildFileSharingPayload(FileSharing fileSharing) {
            ushort? passwordHash = ParseLegacyHashOrThrow(GetFileSharingReservationPassword(fileSharing), "write-reservation password");
            using var stream = new MemoryStream();
            WriteUInt16(stream, fileSharing.ReadOnlyRecommended?.Value == true ? (ushort)1 : (ushort)0);
            WriteUInt16(stream, passwordHash ?? 0);
            if (passwordHash.HasValue) {
                string userName = fileSharing.UserName?.Value ?? string.Empty;
                if (userName.Length > 54) {
                    throw new NotSupportedException("Native XLS saving supports write-reservation user names up to 54 characters.");
                }

                byte[] userNamePayload = BuildUnicodeStringPayload(userName);
                stream.Write(userNamePayload, 0, userNamePayload.Length);
            } else {
                WriteUInt16(stream, 0);
            }

            return stream.ToArray();
        }

        private static byte[] BuildWriteAccessPayload(string userName) {
            if (userName.Length > 54) {
                throw new NotSupportedException("Native XLS saving supports write-reservation user names up to 54 characters.");
            }

            byte[] userNamePayload = BuildUnicodeStringPayload(userName);
            if (userNamePayload.Length > 112) {
                throw new NotSupportedException("Native XLS saving supports write-reservation user names up to 54 characters.");
            }

            byte[] payload = Enumerable.Repeat((byte)0x20, 112).ToArray();
            Buffer.BlockCopy(userNamePayload, 0, payload, 0, userNamePayload.Length);
            return payload;
        }

        private static string? GetFileSharingReservationPassword(FileSharing fileSharing) {
            foreach (OpenXmlAttribute attribute in fileSharing.GetAttributes()) {
                if (string.Equals(attribute.LocalName, "reservationPassword", StringComparison.Ordinal)) {
                    return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
                }
            }

            return null;
        }

        private static LegacyXlsWorksheetProtection ExtractWorksheetProtection(ExcelSheet sheet) {
            SheetProtection? protection = sheet.WorksheetPart.Worksheet?.Elements<SheetProtection>().FirstOrDefault();
            if (protection == null) {
                return LegacyXlsWorksheetProtection.None;
            }

            return new LegacyXlsWorksheetProtection(
                protection.Sheet?.Value != false,
                ParseLegacyHashOrThrow(protection.Password?.Value, "worksheet protection password"),
                protection.Objects?.Value,
                protection.Scenarios?.Value);
        }

        private static void WriteWorksheetProtectionRecords(Stream stream, LegacyXlsWorksheetProtection protection) {
            if (!protection.IsProtected) {
                return;
            }

            WriteRecord(stream, 0x0012, BuildUInt16Payload(1));
            if (protection.PasswordHash.HasValue) {
                WriteRecord(stream, 0x0013, BuildUInt16Payload(protection.PasswordHash.Value));
            }

            if (protection.ProtectObjects.HasValue) {
                WriteRecord(stream, 0x0063, BuildUInt16Payload(protection.ProtectObjects.Value ? (ushort)1 : (ushort)0));
            }

            if (protection.ProtectScenarios.HasValue) {
                WriteRecord(stream, 0x00dd, BuildUInt16Payload(protection.ProtectScenarios.Value ? (ushort)1 : (ushort)0));
            }
        }

        private static ushort? ParseLegacyHashOrThrow(string? value, string feature) {
            if (!string.IsNullOrWhiteSpace(value)
                && ushort.TryParse(value!.Trim(), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out ushort hash)) {
                return hash;
            }

            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            throw new NotSupportedException($"Native XLS saving does not support invalid {feature} hash '{value}'. Save as .xlsx or remove this protection before saving as .xls.");
        }

        private static byte[] BuildDoublePayload(double value) {
            byte[] bytes = BitConverter.GetBytes(value);
            if (!BitConverter.IsLittleEndian) {
                Array.Reverse(bytes);
            }

            return bytes;
        }

        private static byte[] BuildSetupPayload(LegacyXlsWorksheetPageSetup pageSetup) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0);
            WriteUInt16(stream, checked((ushort)(pageSetup.Scale ?? 100U)));
            WriteUInt16(stream, 1);
            WriteUInt16(stream, checked((ushort)(pageSetup.FitToWidth ?? 1U)));
            WriteUInt16(stream, checked((ushort)(pageSetup.FitToHeight ?? 1U)));

            ushort flags = 0;
            if (pageSetup.PageOrder == ExcelPageOrder.OverThenDown) {
                flags |= 0x0001;
            }

            if (pageSetup.Orientation == ExcelPageOrientation.Portrait) {
                flags |= 0x0002;
            }

            WriteUInt16(stream, flags);
            WriteUInt16(stream, 300);
            WriteUInt16(stream, 300);
            WriteDouble(stream, pageSetup.HeaderMargin ?? 0.3d);
            WriteDouble(stream, pageSetup.FooterMargin ?? 0.3d);
            WriteUInt16(stream, 1);
            return stream.ToArray();
        }

        private static byte[] BuildPrinterSettingsPayload(byte[] printerSettingsBytes) {
            if (printerSettingsBytes.Length > ushort.MaxValue - 2) {
                throw new NotSupportedException("Native XLS saving supports BIFF printer settings payloads up to 65,533 bytes.");
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, 0);
            stream.Write(printerSettingsBytes, 0, printerSettingsBytes.Length);
            return stream.ToArray();
        }

        private static byte[] BuildUnicodeStringPayload(string text) {
            byte[] textBytes = EncodeUnicodeString(text, out byte flags);
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)text.Length));
            stream.WriteByte(flags);
            stream.Write(textBytes, 0, textBytes.Length);
            return stream.ToArray();
        }

        private static byte[] BuildHeaderFooterExtensionPayload(LegacyXlsWorksheetPageSetup pageSetup) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x089c);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
            stream.Write(new byte[16], 0, 16);

            ushort flags = 0;
            if (pageSetup.DifferentOddEvenHeaderFooter) {
                flags |= 0x0001;
            }

            if (pageSetup.DifferentFirstHeaderFooter) {
                flags |= 0x0002;
            }

            if (pageSetup.ScaleHeaderFooterWithDocument) {
                flags |= 0x0004;
            }

            if (pageSetup.AlignHeaderFooterWithMargins) {
                flags |= 0x0008;
            }

            string evenHeaderText = pageSetup.EvenHeaderText ?? string.Empty;
            string evenFooterText = pageSetup.EvenFooterText ?? string.Empty;
            string firstHeaderText = pageSetup.FirstHeaderText ?? string.Empty;
            string firstFooterText = pageSetup.FirstFooterText ?? string.Empty;

            WriteUInt16(stream, flags);
            WriteHeaderFooterExtensionLength(stream, evenHeaderText);
            WriteHeaderFooterExtensionLength(stream, evenFooterText);
            WriteHeaderFooterExtensionLength(stream, firstHeaderText);
            WriteHeaderFooterExtensionLength(stream, firstFooterText);
            WriteHeaderFooterExtensionStringBody(stream, evenHeaderText);
            WriteHeaderFooterExtensionStringBody(stream, evenFooterText);
            WriteHeaderFooterExtensionStringBody(stream, firstHeaderText);
            WriteHeaderFooterExtensionStringBody(stream, firstFooterText);
            return stream.ToArray();
        }

        private static void WriteHeaderFooterExtensionLength(Stream stream, string value) {
            if (value.Length > ushort.MaxValue) {
                throw new NotSupportedException("Native XLS saving supports header and footer variant text up to 65535 characters.");
            }

            WriteUInt16(stream, checked((ushort)value.Length));
        }

        private static void WriteHeaderFooterExtensionStringBody(Stream stream, string value) {
            if (value.Length == 0) {
                return;
            }

            byte[] textBytes = EncodeUnicodeString(value, out byte flags);
            stream.WriteByte(flags);
            stream.Write(textBytes, 0, textBytes.Length);
        }

        private static byte[] BuildCodeNamePayload(string codeName) {
            string trimmed = codeName.Trim();
            if (trimmed.Length == 0) {
                return Array.Empty<byte>();
            }

            if (trimmed.Length > 31) {
                throw new NotSupportedException($"Native XLS saving supports VBA object code names up to 31 characters; this workbook uses '{trimmed}'.");
            }

            return BuildUnicodeStringPayload(trimmed);
        }

        private static byte[] BuildSheetTabIdsPayload(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets) {
            DocumentFormat.OpenXml.Spreadsheet.Sheet[] sheetElements =
                document.WorkbookRoot.Sheets?.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToArray()
                ?? Array.Empty<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
            using var stream = new MemoryStream();
            for (int i = 0; i < sheets.Count; i++) {
                uint sheetId = i < sheetElements.Length ? sheetElements[i].SheetId?.Value ?? 0U : 0U;
                ushort tabId = sheetId > 0U && sheetId <= ushort.MaxValue
                    ? checked((ushort)sheetId)
                    : checked((ushort)(i + 1));
                WriteUInt16(stream, tabId);
            }

            return stream.ToArray();
        }

        private static byte[] BuildHorizontalPageBreaksPayload(IReadOnlyList<int> rowPageBreaks) {
            if (rowPageBreaks.Count > 1026) {
                throw new NotSupportedException("Native XLS saving supports up to 1,026 manual row page breaks per worksheet.");
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)rowPageBreaks.Count));
            foreach (int row in rowPageBreaks) {
                if (row <= 0 || row > ushort.MaxValue) {
                    throw new NotSupportedException("Native XLS saving supports row page breaks within the BIFF8 encodable worksheet row break limit of 65,535.");
                }

                WriteUInt16(stream, checked((ushort)row));
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 255);
            }

            return stream.ToArray();
        }

        private static byte[] BuildVerticalPageBreaksPayload(IReadOnlyList<int> columnPageBreaks) {
            if (columnPageBreaks.Count > 1026) {
                throw new NotSupportedException("Native XLS saving supports up to 1,026 manual column page breaks per worksheet.");
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)columnPageBreaks.Count));
            foreach (int column in columnPageBreaks) {
                if (column <= 0 || column > 256) {
                    throw new NotSupportedException("Native XLS saving supports column page breaks within the BIFF8 worksheet limit of 256 columns.");
                }

                WriteUInt16(stream, checked((ushort)column));
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 65535);
            }

            return stream.ToArray();
        }

        private static string? BuildHeaderFooterText(string? left, string? center, string? right) {
            var builder = new StringBuilder();
            if (!string.IsNullOrEmpty(left)) {
                builder.Append("&L").Append(EscapeHeaderFooterText(left!));
            }

            if (!string.IsNullOrEmpty(center)) {
                builder.Append("&C").Append(EscapeHeaderFooterText(center!));
            }

            if (!string.IsNullOrEmpty(right)) {
                builder.Append("&R").Append(EscapeHeaderFooterText(right!));
            }

            return builder.Length == 0 ? null : builder.ToString();
        }

        private static string EscapeHeaderFooterText(string text) {
            var builder = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch == '&' && (i + 1 >= text.Length || !IsHeaderFooterTokenStarter(text[i + 1]))) {
                    builder.Append("&&");
                } else {
                    builder.Append(ch);
                }
            }

            return builder.ToString();
        }

        private static bool IsHeaderFooterTokenStarter(char ch) {
            if (ch >= '0' && ch <= '9') {
                return true;
            }

            return ch == '"'
                || ch == '&'
                || ch == 'A'
                || ch == 'B'
                || ch == 'D'
                || ch == 'E'
                || ch == 'F'
                || ch == 'G'
                || ch == 'I'
                || ch == 'K'
                || ch == 'L'
                || ch == 'N'
                || ch == 'P'
                || ch == 'R'
                || ch == 'S'
                || ch == 'T'
                || ch == 'U'
                || ch == 'X'
                || ch == 'Y'
                || ch == 'Z';
        }

        private static double? NormalizeDefaultColumnWidth(double? width) {
            if (!width.HasValue) {
                return null;
            }

            double rounded = Math.Round(width.Value);
            return rounded > 0 && rounded <= 255 ? rounded : null;
        }

        private static byte[] BuildUInt16Payload(ushort value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, value);
            return stream.ToArray();
        }

        private static byte[] EncodeShortUnicodeString(string text, out byte flags) {
            return EncodeUnicodeString(text, out flags);
        }

        private static byte[] EncodeUnicodeString(string text, out byte flags) {
            if (CanUseCompressedString(text)) {
                flags = 0;
                return Encoding.ASCII.GetBytes(text);
            }

            flags = 1;
            return Encoding.Unicode.GetBytes(text);
        }

        private static bool CanUseCompressedString(string text) {
            for (int i = 0; i < text.Length; i++) {
                if (text[i] > 0x7f) {
                    return false;
                }
            }

            return true;
        }

        private static long GetEncodedStringByteCount(string text) {
            return CanUseCompressedString(text)
                ? text.Length
                : (long)text.Length * 2L;
        }

        private static void WriteRecord(Stream stream, ushort type, byte[] payload) {
            if (payload.Length > ushort.MaxValue) {
                throw new NotSupportedException($"Native XLS saving does not yet support BIFF record 0x{type:X4} payload lengths outside BIFF8 limits. Save as .xlsx or remove this feature before saving as .xls.");
            }

            WriteUInt16(stream, type);
            WriteUInt16(stream, (ushort)payload.Length);
            stream.Write(payload, 0, payload.Length);
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
            stream.WriteByte((byte)((value >> 16) & 0xff));
            stream.WriteByte((byte)((value >> 24) & 0xff));
        }

        private static void WriteDouble(Stream stream, double value) {
            byte[] bytes = BitConverter.GetBytes(value);
            if (!BitConverter.IsLittleEndian) {
                Array.Reverse(bytes);
            }

            stream.Write(bytes, 0, bytes.Length);
        }

        private static void WriteUInt16(byte[] buffer, int offset, ushort value) {
            buffer[offset] = (byte)(value & 0xff);
            buffer[offset + 1] = (byte)((value >> 8) & 0xff);
        }

        private static ushort ReadUInt16(byte[] buffer, int offset) {
            return (ushort)(buffer[offset] | (buffer[offset + 1] << 8));
        }

        private static void WriteUInt32(byte[] buffer, int offset, uint value) {
            buffer[offset] = (byte)(value & 0xff);
            buffer[offset + 1] = (byte)((value >> 8) & 0xff);
            buffer[offset + 2] = (byte)((value >> 16) & 0xff);
            buffer[offset + 3] = (byte)((value >> 24) & 0xff);
        }

        private static void WriteFormulaFollowUpRecords(Stream stream, LegacyXlsCell cell) {
            if (cell.ArrayFormulaPayload.Length > 0) {
                WriteRecord(stream, 0x0221, cell.ArrayFormulaPayload);
            }

            if (cell.SharedFormulaPayload.Length > 0) {
                WriteRecord(stream, 0x04bc, cell.SharedFormulaPayload);
            }
        }

        private sealed class LegacyXlsArrayFormulaTable {
            private readonly List<LegacyXlsArrayFormulaDefinition> _definitions;

            private LegacyXlsArrayFormulaTable(List<LegacyXlsArrayFormulaDefinition> definitions) {
                _definitions = definitions;
            }

            internal static LegacyXlsArrayFormulaTable Create(ExcelSheet sheet, int sheetIndex, LegacyXlsFormulaNameIndex formulaNameIndex) {
                SheetData? sheetData = sheet.WorksheetPart.Worksheet?.GetFirstChild<SheetData>();
                var definitions = new List<LegacyXlsArrayFormulaDefinition>();
                if (sheetData == null) {
                    return new LegacyXlsArrayFormulaTable(definitions);
                }

                Dictionary<ulong, Cell> cellsByAddress = BuildCellAddressMap(sheetData);
                foreach (KeyValuePair<ulong, Cell> entry in cellsByAddress) {
                    Cell cell = entry.Value;
                    CellFormula? formula = cell.CellFormula;
                    if (formula?.FormulaType?.Value != DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Array) {
                        continue;
                    }

                    DecodeCellKey(entry.Key, out ushort anchorRow, out ushort anchorColumn);
                    if (!TryParseFormulaRange(formula.Reference?.Value, out ushort firstRow, out ushort firstColumn, out ushort lastRow, out ushort lastColumn)
                        || anchorRow < firstRow
                        || anchorRow > lastRow
                        || anchorColumn < firstColumn
                        || anchorColumn > lastColumn) {
                        throw new NotSupportedException($"Native XLS saving does not support array formula range '{formula.Reference?.Value}' at {ToA1Address(anchorRow, anchorColumn)}. Save as .xlsx or remove this formula before saving as .xls.");
                    }

                    if (!LegacyXlsFormulaEncoder.TryEncode(formula.Text ?? string.Empty, formulaNameIndex, sheetIndex, out byte[] formulaTokens, out string? reason)) {
                        throw new NotSupportedException($"Native XLS saving does not yet support array formula '{formula.Text}' at {ToA1Address(anchorRow, anchorColumn)}: {reason} Save as .xlsx or remove this formula before saving as .xls.");
                    }

                    EnsureSupportedArrayFormulaPayloadLength(formulaTokens, ToA1Address(anchorRow, anchorColumn));
                    var definition = new LegacyXlsArrayFormulaDefinition(
                        anchorRow,
                        anchorColumn,
                        firstRow,
                        firstColumn,
                        lastRow,
                        lastColumn,
                        BuildArrayFormulaPayload(firstRow, firstColumn, lastRow, lastColumn, formulaTokens));
                    EnsureArrayFormulaCachedCells(cellsByAddress, definition);
                    definitions.Add(definition);
                }

                return new LegacyXlsArrayFormulaTable(definitions);
            }

            internal bool TryGetDefinition(ushort row, ushort column, out LegacyXlsArrayFormulaDefinition definition) {
                foreach (LegacyXlsArrayFormulaDefinition candidate in _definitions) {
                    if (candidate.Contains(row, column)) {
                        definition = candidate;
                        return true;
                    }
                }

                definition = default;
                return false;
            }

            private static Dictionary<ulong, Cell> BuildCellAddressMap(SheetData sheetData) {
                var cells = new Dictionary<ulong, Cell>();
                foreach (Row row in sheetData.Elements<Row>()) {
                    uint rowIndex = row.RowIndex?.Value ?? 0U;
                    int sequentialColumn = 1;
                    foreach (Cell cell in row.Elements<Cell>()) {
                        uint effectiveRow = rowIndex;
                        int effectiveColumn = sequentialColumn;
                        if (!string.IsNullOrEmpty(cell.CellReference?.Value)) {
                            ParseCellReference(cell.CellReference!.Value!, out effectiveRow, out effectiveColumn);
                        }

                        if (effectiveRow > 0 && effectiveColumn > 0 && effectiveRow <= 65536U && effectiveColumn <= 256) {
                            cells[EncodeCellKey(checked((ushort)(effectiveRow - 1U)), checked((ushort)(effectiveColumn - 1)))] = cell;
                        }

                        sequentialColumn = effectiveColumn + 1;
                    }
                }

                return cells;
            }

            private static void EnsureArrayFormulaCachedCells(IReadOnlyDictionary<ulong, Cell> cellsByAddress, LegacyXlsArrayFormulaDefinition definition) {
                for (int row = definition.FirstRow; row <= definition.LastRow; row++) {
                    for (int column = definition.FirstColumn; column <= definition.LastColumn; column++) {
                        ushort currentRow = checked((ushort)row);
                        ushort currentColumn = checked((ushort)column);
                        if (!cellsByAddress.TryGetValue(EncodeCellKey(currentRow, currentColumn), out Cell? cell)
                            || string.IsNullOrWhiteSpace(cell.CellValue?.InnerText)) {
                            throw new NotSupportedException($"Native XLS saving requires cached results for every cell in multi-cell array formula range {definition.Range}. Missing cached result at {ToA1Address(currentRow, currentColumn)}.");
                        }

                        if ((currentRow != definition.AnchorRow || currentColumn != definition.AnchorColumn) && cell.CellFormula != null) {
                            throw new NotSupportedException($"Native XLS saving does not support non-anchor formulas inside array formula range {definition.Range}. Remove the formula at {ToA1Address(currentRow, currentColumn)} or save as .xlsx.");
                        }
                    }
                }
            }

            private static ulong EncodeCellKey(ushort row, ushort column) {
                return ((ulong)row << 16) | column;
            }

            private static void DecodeCellKey(ulong key, out ushort row, out ushort column) {
                row = (ushort)(key >> 16);
                column = (ushort)(key & 0xffff);
            }
        }

        private readonly struct LegacyXlsArrayFormulaDefinition {
            internal LegacyXlsArrayFormulaDefinition(ushort anchorRow, ushort anchorColumn, ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn, byte[] payload) {
                AnchorRow = anchorRow;
                AnchorColumn = anchorColumn;
                FirstRow = firstRow;
                FirstColumn = firstColumn;
                LastRow = lastRow;
                LastColumn = lastColumn;
                Payload = payload;
            }

            internal ushort AnchorRow { get; }

            internal ushort AnchorColumn { get; }

            internal ushort FirstRow { get; }

            internal ushort FirstColumn { get; }

            internal ushort LastRow { get; }

            internal ushort LastColumn { get; }

            internal byte[] Payload { get; }

            internal string Range => ToA1Address(FirstRow, FirstColumn) + ":" + ToA1Address(LastRow, LastColumn);

            internal bool Contains(ushort row, ushort column) {
                return row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;
            }
        }

        private sealed class LegacyXlsSharedFormulaTable {
            private readonly Dictionary<uint, LegacyXlsSharedFormulaDefinition> _definitions;

            private LegacyXlsSharedFormulaTable(Dictionary<uint, LegacyXlsSharedFormulaDefinition> definitions) {
                _definitions = definitions;
            }

            internal static LegacyXlsSharedFormulaTable Create(ExcelSheet sheet, int sheetIndex, LegacyXlsFormulaNameIndex formulaNameIndex) {
                SheetData? sheetData = sheet.WorksheetPart.Worksheet?.GetFirstChild<SheetData>();
                var definitions = new Dictionary<uint, LegacyXlsSharedFormulaDefinition>();
                if (sheetData == null) {
                    return new LegacyXlsSharedFormulaTable(definitions);
                }

                foreach (Row row in sheetData.Elements<Row>()) {
                    uint rowIndex = row.RowIndex?.Value ?? 0U;
                    int sequentialColumn = 1;
                    foreach (Cell cell in row.Elements<Cell>()) {
                        CellFormula? formula = cell.CellFormula;
                        if (formula?.FormulaType?.Value != DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared) {
                            if (!string.IsNullOrEmpty(cell.CellReference?.Value)) {
                                ParseCellReference(cell.CellReference!.Value!, out _, out sequentialColumn);
                            }

                            sequentialColumn++;
                            continue;
                        }

                        uint effectiveRow = rowIndex;
                        int effectiveColumn = sequentialColumn;
                        if (!string.IsNullOrEmpty(cell.CellReference?.Value)) {
                            ParseCellReference(cell.CellReference!.Value!, out effectiveRow, out effectiveColumn);
                        }

                        sequentialColumn = effectiveColumn + 1;
                        if (effectiveRow == 0 || effectiveColumn <= 0 || formula.SharedIndex == null || string.IsNullOrWhiteSpace(formula.Text)) {
                            continue;
                        }

                        uint sharedIndex = formula.SharedIndex.Value;
                        if (definitions.ContainsKey(sharedIndex)) {
                            throw new NotSupportedException($"Native XLS saving found duplicate shared formula definition index {sharedIndex} on worksheet '{sheet.Name}'. Save as .xlsx or remove this formula before saving as .xls.");
                        }

                        ushort anchorRow = checked((ushort)(effectiveRow - 1U));
                        ushort anchorColumn = checked((ushort)(effectiveColumn - 1));
                        if (!TryParseFormulaRange(formula.Reference?.Value, out ushort firstRow, out ushort firstColumn, out ushort lastRow, out ushort lastColumn)
                            || anchorRow < firstRow
                            || anchorRow > lastRow
                            || anchorColumn < firstColumn
                            || anchorColumn > lastColumn) {
                            throw new NotSupportedException($"Native XLS saving does not support shared formula range '{formula.Reference?.Value}' at {ToA1Address(anchorRow, anchorColumn)}. Save as .xlsx or remove this formula before saving as .xls.");
                        }

                        if (!LegacyXlsFormulaEncoder.TryEncode(formula.Text, formulaNameIndex, sheetIndex, out byte[] formulaTokens, out string? reason)) {
                            throw new NotSupportedException($"Native XLS saving does not yet support shared formula '{formula.Text}' at {ToA1Address(anchorRow, anchorColumn)}: {reason} Save as .xlsx or remove this formula before saving as .xls.");
                        }

                        if (!TryBuildSharedFormulaTokens(formulaTokens, anchorRow, anchorColumn, out byte[] sharedFormulaTokens)) {
                            throw new NotSupportedException($"Native XLS saving does not yet support shared formula token conversion for '{formula.Text}' at {ToA1Address(anchorRow, anchorColumn)}. Save as .xlsx or remove this formula before saving as .xls.");
                        }

                        EnsureSupportedSharedFormulaPayloadLength(sharedFormulaTokens, ToA1Address(anchorRow, anchorColumn));
                        definitions[sharedIndex] = new LegacyXlsSharedFormulaDefinition(
                            anchorRow,
                            anchorColumn,
                            formula.Text,
                            BuildSharedFormulaPayload(firstRow, firstColumn, lastRow, lastColumn, sharedFormulaTokens));
                    }
                }

                return new LegacyXlsSharedFormulaTable(definitions);
            }

            internal bool TryGetDefinition(uint? sharedIndex, out LegacyXlsSharedFormulaDefinition definition) {
                definition = default;
                return sharedIndex.HasValue && _definitions.TryGetValue(sharedIndex.Value, out definition);
            }
        }

        private readonly struct LegacyXlsSharedFormulaDefinition {
            internal LegacyXlsSharedFormulaDefinition(ushort anchorRow, ushort anchorColumn, string formulaText, byte[] payload) {
                AnchorRow = anchorRow;
                AnchorColumn = anchorColumn;
                FormulaText = formulaText;
                Payload = payload;
            }

            internal ushort AnchorRow { get; }

            internal ushort AnchorColumn { get; }

            internal string FormulaText { get; }

            internal byte[] Payload { get; }
        }

        private readonly struct LegacyXlsCell {
            private LegacyXlsCell(
                ushort row,
                ushort column,
                ushort styleIndex,
                LegacyXlsCellKind kind,
                string? textValue,
                double numberValue,
                bool booleanValue,
                byte errorValue,
                byte[]? formulaTokens,
                byte[]? formulaExtraData,
                byte[]? arrayFormulaPayload,
                byte[]? sharedFormulaPayload,
                IReadOnlyList<LegacyXlsTextFormattingRun>? textFormattingRuns) {
                Row = row;
                Column = column;
                StyleIndex = styleIndex;
                Kind = kind;
                TextValue = textValue;
                NumberValue = numberValue;
                BooleanValue = booleanValue;
                ErrorValue = errorValue;
                FormulaTokens = formulaTokens ?? Array.Empty<byte>();
                FormulaExtraData = formulaExtraData ?? Array.Empty<byte>();
                ArrayFormulaPayload = arrayFormulaPayload ?? Array.Empty<byte>();
                SharedFormulaPayload = sharedFormulaPayload ?? Array.Empty<byte>();
                TextFormattingRuns = textFormattingRuns ?? Array.Empty<LegacyXlsTextFormattingRun>();
            }

            internal ushort Row { get; }

            internal ushort Column { get; }

            internal ushort StyleIndex { get; }

            internal LegacyXlsCellKind Kind { get; }

            internal string? TextValue { get; }

            internal double NumberValue { get; }

            internal bool BooleanValue { get; }

            internal byte ErrorValue { get; }

            internal byte[] FormulaTokens { get; }

            internal byte[] FormulaExtraData { get; }

            internal byte[] ArrayFormulaPayload { get; }

            internal byte[] SharedFormulaPayload { get; }

            internal IReadOnlyList<LegacyXlsTextFormattingRun> TextFormattingRuns { get; }

            internal static LegacyXlsCell Blank(ushort row, ushort column, ushort styleIndex) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.Blank, null, 0, false, 0, null, null, null, null, null);

            internal static LegacyXlsCell Text(ushort row, ushort column, ushort styleIndex, string value, IReadOnlyList<LegacyXlsTextFormattingRun>? textFormattingRuns = null) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.Text, value, 0, false, 0, null, null, null, null, textFormattingRuns);

            internal static LegacyXlsCell Number(ushort row, ushort column, ushort styleIndex, double value) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.Number, null, value, false, 0, null, null, null, null, null);

            internal static LegacyXlsCell Boolean(ushort row, ushort column, ushort styleIndex, bool value) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.Boolean, null, 0, value, 0, null, null, null, null, null);

            internal static LegacyXlsCell Error(ushort row, ushort column, ushort styleIndex, byte errorValue) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.Error, null, 0, false, errorValue, null, null, null, null, null);

            internal static LegacyXlsCell FormulaNumber(ushort row, ushort column, ushort styleIndex, double value, byte[] formulaTokens, byte[]? formulaExtraData = null, byte[]? arrayFormulaPayload = null, byte[]? sharedFormulaPayload = null) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.FormulaNumber, null, value, false, 0, formulaTokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload, null);

            internal static LegacyXlsCell FormulaBoolean(ushort row, ushort column, ushort styleIndex, bool value, byte[] formulaTokens, byte[]? formulaExtraData = null, byte[]? arrayFormulaPayload = null, byte[]? sharedFormulaPayload = null) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.FormulaBoolean, null, 0, value, 0, formulaTokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload, null);

            internal static LegacyXlsCell FormulaText(ushort row, ushort column, ushort styleIndex, string value, byte[] formulaTokens, byte[]? formulaExtraData = null, byte[]? arrayFormulaPayload = null, byte[]? sharedFormulaPayload = null) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.FormulaText, value, 0, false, 0, formulaTokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload, null);

            internal static LegacyXlsCell FormulaError(ushort row, ushort column, ushort styleIndex, byte errorValue, byte[] formulaTokens, byte[]? formulaExtraData = null, byte[]? arrayFormulaPayload = null, byte[]? sharedFormulaPayload = null) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.FormulaError, null, 0, false, errorValue, formulaTokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload, null);
        }

        private enum LegacyXlsCellKind {
            Blank,
            Text,
            Number,
            Boolean,
            Error,
            FormulaNumber,
            FormulaBoolean,
            FormulaText,
            FormulaError
        }

        private sealed class LegacyXlsWorksheetLayout {
            internal LegacyXlsWorksheetLayout(
                double? defaultColumnWidth,
                double? defaultRowHeight,
                bool defaultRowsHidden,
                IReadOnlyList<LegacyXlsColumnLayout> columns,
                IReadOnlyList<LegacyXlsRowLayout> rows,
                IReadOnlyList<LegacyXlsMergedRange> mergedRanges,
                IReadOnlyList<LegacyXlsWorksheetView> views,
                LegacyXlsWorksheetPageSetup pageSetup,
                LegacyXlsWorksheetProtection protection) {
                DefaultColumnWidth = defaultColumnWidth;
                DefaultRowHeight = defaultRowHeight;
                DefaultRowsHidden = defaultRowsHidden;
                Columns = columns;
                Rows = rows;
                MergedRanges = mergedRanges;
                Views = views ?? throw new ArgumentNullException(nameof(views));
                PageSetup = pageSetup;
                Protection = protection;
            }

            internal double? DefaultColumnWidth { get; }

            internal double? DefaultRowHeight { get; }

            internal bool DefaultRowsHidden { get; }

            internal IReadOnlyList<LegacyXlsColumnLayout> Columns { get; }

            internal IReadOnlyList<LegacyXlsRowLayout> Rows { get; }

            internal IReadOnlyList<LegacyXlsMergedRange> MergedRanges { get; }

            internal IReadOnlyList<LegacyXlsWorksheetView> Views { get; }

            internal LegacyXlsWorksheetView View => Views.Count == 0
                ? throw new InvalidOperationException("A worksheet layout must contain at least one view.")
                : Views[0];

            internal LegacyXlsWorksheetPageSetup PageSetup { get; }

            internal LegacyXlsWorksheetProtection Protection { get; }
        }

        private readonly struct LegacyXlsColumnLayout {
            internal LegacyXlsColumnLayout(ushort firstColumn, ushort lastColumn, double? width, bool hidden, ushort styleIndex, byte outlineLevel, bool collapsed) {
                FirstColumn = firstColumn;
                LastColumn = lastColumn;
                Width = width;
                Hidden = hidden;
                StyleIndex = styleIndex;
                OutlineLevel = outlineLevel;
                Collapsed = collapsed;
            }

            internal ushort FirstColumn { get; }

            internal ushort LastColumn { get; }

            internal double? Width { get; }

            internal bool Hidden { get; }

            internal ushort StyleIndex { get; }

            internal byte OutlineLevel { get; }

            internal bool Collapsed { get; }
        }

        private readonly struct LegacyXlsRowLayout {
            internal LegacyXlsRowLayout(ushort row, double? height, bool hidden, bool customHeight, bool customFormat, ushort styleIndex, byte outlineLevel, bool collapsed) {
                Row = row;
                Height = height;
                Hidden = hidden;
                CustomHeight = customHeight;
                CustomFormat = customFormat;
                StyleIndex = styleIndex;
                OutlineLevel = outlineLevel;
                Collapsed = collapsed;
            }

            internal ushort Row { get; }

            internal double? Height { get; }

            internal bool Hidden { get; }

            internal bool CustomHeight { get; }

            internal bool CustomFormat { get; }

            internal ushort StyleIndex { get; }

            internal byte OutlineLevel { get; }

            internal bool Collapsed { get; }
        }

        private readonly struct LegacyXlsMergedRange {
            internal LegacyXlsMergedRange(ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                FirstRow = firstRow;
                FirstColumn = firstColumn;
                LastRow = lastRow;
                LastColumn = lastColumn;
            }

            internal ushort FirstRow { get; }

            internal ushort FirstColumn { get; }

            internal ushort LastRow { get; }

            internal ushort LastColumn { get; }
        }

        private readonly struct LegacyXlsWorksheetView {
            internal LegacyXlsWorksheetView(int frozenRowCount, int frozenColumnCount, bool showFormulas, bool showGridlines, bool showRowColumnHeadings, bool showZeroValues, bool rightToLeft, bool defaultGridColor, ushort gridLineColorIndex, bool showOutlineSymbols, bool tabSelected, bool pageBreakPreview, bool pageLayoutView, bool frozenWithoutSplit, LegacyXlsSplitPaneView? splitPane, uint? zoomScale, uint? zoomScaleNormal, LegacyXlsWindowTopLeftCell topLeftCell, IReadOnlyList<LegacyXlsSelection> selections) {
                FrozenRowCount = frozenRowCount;
                FrozenColumnCount = frozenColumnCount;
                ShowFormulas = showFormulas;
                ShowGridlines = showGridlines;
                ShowRowColumnHeadings = showRowColumnHeadings;
                ShowZeroValues = showZeroValues;
                RightToLeft = rightToLeft;
                DefaultGridColor = defaultGridColor;
                GridLineColorIndex = gridLineColorIndex;
                ShowOutlineSymbols = showOutlineSymbols;
                TabSelected = tabSelected;
                PageBreakPreview = pageBreakPreview;
                PageLayoutView = pageLayoutView;
                FrozenWithoutSplit = frozenWithoutSplit;
                SplitPane = splitPane;
                ZoomScale = zoomScale;
                ZoomScaleNormal = zoomScaleNormal;
                TopLeftCell = topLeftCell;
                Selections = selections ?? throw new ArgumentNullException(nameof(selections));
            }

            internal int FrozenRowCount { get; }

            internal int FrozenColumnCount { get; }

            internal bool ShowFormulas { get; }

            internal bool ShowGridlines { get; }

            internal bool ShowRowColumnHeadings { get; }

            internal bool ShowZeroValues { get; }

            internal bool RightToLeft { get; }

            internal bool DefaultGridColor { get; }

            internal ushort GridLineColorIndex { get; }

            internal bool ShowOutlineSymbols { get; }

            internal bool TabSelected { get; }

            internal bool PageBreakPreview { get; }

            internal bool PageLayoutView { get; }

            internal bool FrozenWithoutSplit { get; }

            internal LegacyXlsSplitPaneView? SplitPane { get; }

            internal uint? ZoomScale { get; }

            internal uint? ZoomScaleNormal { get; }

            internal LegacyXlsWindowTopLeftCell TopLeftCell { get; }

            internal IReadOnlyList<LegacyXlsSelection> Selections { get; }
        }

        private readonly struct LegacyXlsSplitPaneView {
            internal LegacyXlsSplitPaneView(ushort horizontalSplit, ushort verticalSplit, ushort topRow, ushort leftColumn, byte activePane) {
                HorizontalSplit = horizontalSplit;
                VerticalSplit = verticalSplit;
                TopRow = topRow;
                LeftColumn = leftColumn;
                ActivePane = activePane;
            }

            internal ushort HorizontalSplit { get; }

            internal ushort VerticalSplit { get; }

            internal ushort TopRow { get; }

            internal ushort LeftColumn { get; }

            internal byte ActivePane { get; }
        }

        private readonly struct LegacyXlsWindowTopLeftCell {
            internal LegacyXlsWindowTopLeftCell(ushort row, ushort column) {
                Row = row;
                Column = column;
            }

            internal ushort Row { get; }

            internal ushort Column { get; }
        }

        private sealed class LegacyXlsWorksheetPageSetup {
            internal LegacyXlsWorksheetPageSetup(
                ExcelPageOrientation? orientation,
                double? leftMargin,
                double? rightMargin,
                double? topMargin,
                double? bottomMargin,
                double? headerMargin,
                double? footerMargin,
                uint? fitToWidth,
                uint? fitToHeight,
                uint? scale,
                ExcelPageOrder? pageOrder,
                bool? fitToPage,
                bool? printGridLines,
                bool? printHeadings,
                bool? horizontalCentered,
                bool? verticalCentered,
                bool? gridLinesSet,
                bool? applyOutlineStyles,
                bool? summaryRowsBelow,
                bool? summaryColumnsRight,
                string? headerText,
                string? footerText,
                string? firstHeaderText,
                string? firstFooterText,
                string? evenHeaderText,
                string? evenFooterText,
                bool differentFirstHeaderFooter,
                bool differentOddEvenHeaderFooter,
                bool scaleHeaderFooterWithDocument,
                bool alignHeaderFooterWithMargins,
                IReadOnlyList<int> rowPageBreaks,
                IReadOnlyList<int> columnPageBreaks,
                byte[]? printerSettingsPayload) {
                Orientation = orientation;
                LeftMargin = leftMargin;
                RightMargin = rightMargin;
                TopMargin = topMargin;
                BottomMargin = bottomMargin;
                HeaderMargin = headerMargin;
                FooterMargin = footerMargin;
                FitToWidth = fitToWidth;
                FitToHeight = fitToHeight;
                Scale = scale;
                PageOrder = pageOrder;
                FitToPage = fitToPage;
                PrintGridLines = printGridLines;
                PrintHeadings = printHeadings;
                HorizontalCentered = horizontalCentered;
                VerticalCentered = verticalCentered;
                GridLinesSet = gridLinesSet;
                ApplyOutlineStyles = applyOutlineStyles;
                SummaryRowsBelow = summaryRowsBelow;
                SummaryColumnsRight = summaryColumnsRight;
                HeaderText = headerText;
                FooterText = footerText;
                FirstHeaderText = firstHeaderText;
                FirstFooterText = firstFooterText;
                EvenHeaderText = evenHeaderText;
                EvenFooterText = evenFooterText;
                DifferentFirstHeaderFooter = differentFirstHeaderFooter;
                DifferentOddEvenHeaderFooter = differentOddEvenHeaderFooter;
                ScaleHeaderFooterWithDocument = scaleHeaderFooterWithDocument;
                AlignHeaderFooterWithMargins = alignHeaderFooterWithMargins;
                RowPageBreaks = rowPageBreaks;
                ColumnPageBreaks = columnPageBreaks;
                PrinterSettingsPayload = printerSettingsPayload;
            }

            internal ExcelPageOrientation? Orientation { get; }

            internal double? LeftMargin { get; }

            internal double? RightMargin { get; }

            internal double? TopMargin { get; }

            internal double? BottomMargin { get; }

            internal double? HeaderMargin { get; }

            internal double? FooterMargin { get; }

            internal uint? FitToWidth { get; }

            internal uint? FitToHeight { get; }

            internal uint? Scale { get; }

            internal ExcelPageOrder? PageOrder { get; }

            internal bool? FitToPage { get; }

            internal bool? PrintGridLines { get; }

            internal bool? PrintHeadings { get; }

            internal bool? HorizontalCentered { get; }

            internal bool? VerticalCentered { get; }

            internal bool? GridLinesSet { get; }

            internal bool? ApplyOutlineStyles { get; }

            internal bool? SummaryRowsBelow { get; }

            internal bool? SummaryColumnsRight { get; }

            internal string? HeaderText { get; }

            internal string? FooterText { get; }

            internal string? FirstHeaderText { get; }

            internal string? FirstFooterText { get; }

            internal string? EvenHeaderText { get; }

            internal string? EvenFooterText { get; }

            internal bool DifferentFirstHeaderFooter { get; }

            internal bool DifferentOddEvenHeaderFooter { get; }

            internal bool ScaleHeaderFooterWithDocument { get; }

            internal bool AlignHeaderFooterWithMargins { get; }

            internal IReadOnlyList<int> RowPageBreaks { get; }

            internal IReadOnlyList<int> ColumnPageBreaks { get; }

            internal byte[]? PrinterSettingsPayload { get; }

            internal bool HasHeaderFooterExtensionRecord =>
                DifferentFirstHeaderFooter
                || DifferentOddEvenHeaderFooter
                || !ScaleHeaderFooterWithDocument
                || !AlignHeaderFooterWithMargins
                || !string.IsNullOrEmpty(FirstHeaderText)
                || !string.IsNullOrEmpty(FirstFooterText)
                || !string.IsNullOrEmpty(EvenHeaderText)
                || !string.IsNullOrEmpty(EvenFooterText);

            internal bool HasSetupRecord =>
                Orientation.HasValue
                || HeaderMargin.HasValue
                || FooterMargin.HasValue
                || FitToWidth.HasValue
                || FitToHeight.HasValue
                || Scale.HasValue
                || PageOrder.HasValue;
        }

        private readonly struct LegacyXlsWorksheetProtection {
            internal LegacyXlsWorksheetProtection(bool isProtected, ushort? passwordHash, bool? protectObjects, bool? protectScenarios) {
                IsProtected = isProtected;
                PasswordHash = passwordHash;
                ProtectObjects = protectObjects;
                ProtectScenarios = protectScenarios;
            }

            internal static LegacyXlsWorksheetProtection None { get; } = new LegacyXlsWorksheetProtection(false, null, null, null);

            internal bool IsProtected { get; }

            internal ushort? PasswordHash { get; }

            internal bool? ProtectObjects { get; }

            internal bool? ProtectScenarios { get; }
        }
    }
}
