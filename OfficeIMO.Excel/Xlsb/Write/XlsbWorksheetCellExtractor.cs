using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;
using System.Globalization;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Creates the bounded cell mutation plan used by the preservation-aware XLSB rewriter.</summary>
    internal static partial class XlsbWorksheetCellExtractor {
        internal static IReadOnlyList<XlsbWriteCell> Extract(
            ExcelDocument document,
            ExcelSheet sheet,
            XlsbSheet sourceSheet,
            int availableCellFormatCount, IXlsbFormulaEncodingContext? formulas = null, XlsbCellTextRuns? textRuns = null) {
            ArgumentNullException.ThrowIfNull(document);
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(sourceSheet);

            ThrowIfUnsupportedWorksheetMutation(sheet, sourceSheet);
            return ExtractCore(document, sheet, sourceSheet, availableCellFormatCount, formulas, textRuns: textRuns);
        }

        internal static IReadOnlyList<XlsbWriteCell> ExtractNew(
            ExcelDocument document,
            ExcelSheet sheet, IXlsbFormulaEncodingContext? formulas = null, XlsbRichTextFontTable? richText = null,
            XlsbCellTextRuns? textRuns = null) {
            ArgumentNullException.ThrowIfNull(document);
            ArgumentNullException.ThrowIfNull(sheet);

            ThrowIfUnsupportedNewWorksheetContent(sheet);
            return ExtractCore(document, sheet, sourceSheet: null, availableCellFormatCount: null, formulas, richText, textRuns);
        }

        private static IReadOnlyList<XlsbWriteCell> ExtractCore(
            ExcelDocument document,
            ExcelSheet sheet,
            XlsbSheet? sourceSheet,
            int? availableCellFormatCount, IXlsbFormulaEncodingContext? formulas, XlsbRichTextFontTable? richText = null,
            XlsbCellTextRuns? textRuns = null) {
            textRuns = richText?.TextRuns ?? textRuns
                ?? XlsbCellTextRuns.ForWorkbook(document);
            var sourceCells = sourceSheet?.Cells.ToDictionary(cell => (cell.Row, cell.Column))
                ?? new Dictionary<(int Row, int Column), XlsbCell>();
            var visitedSourceCells = new HashSet<(int Row, int Column)>();
            var result = new List<XlsbWriteCell>();
            IReadOnlyDictionary<string, string> resolvedFormulaTexts = sheet.BuildResolvedFormulaTextMap();
            OfficeXmlNode sheetData = XlsbWorksheetMarkup.Child(XlsbWorksheetMarkup.Root(sheet), "sheetData");
            if (sheetData.IsNone) return result;

            foreach (OfficeXmlNode row in WorksheetChildren(sheetData, "row")) {
                uint sequentialRow = XlsbWorkbookMarkup.ReadUInt32(row, "r") ?? 0U;
                int sequentialColumn = 1;
                foreach (OfficeXmlNode node in WorksheetChildren(row, "c")) {
                    var cell = new CellMarkup(node);
                    ParseCellReference(cell.Reference, sequentialRow, sequentialColumn, out int cellRow, out int cellColumn);
                    if (cellRow <= 0 || cellRow > 1_048_576 || cellColumn <= 0 || cellColumn > 16_384) {
                        throw new NotSupportedException($"Native XLSB rewriting cannot encode invalid cell reference '{cell.Reference}'.");
                    }

                    sourceCells.TryGetValue((cellRow, cellColumn), out XlsbCell? sourceCell);
                    if (sourceCell != null) visitedSourceCells.Add((cellRow, cellColumn));
                    uint styleIndex = ResolveStyleIndex(
                        cell,
                        sourceCell,
                        sheet.Name,
                        cellRow,
                        cellColumn,
                        allowNewStyles: sourceSheet == null,
                        availableCellFormatCount);
                    if (sourceCell != null
                        && CellMatchesSource(sheet, cell, sourceCell, resolvedFormulaTexts, textRuns)) {
                        var preserved = XlsbWriteCell.PreserveSource(sourceCell, styleIndex);
                        result.Add(formulas?.RebindSourceCell(preserved) ?? preserved);
                        sequentialColumn = cellColumn + 1;
                        continue;
                    }

                    XlsbWriteCell? writeCell = ConvertCell(
                        document,
                        sheet,
                        cell,
                        sourceCell,
                        cellRow,
                        cellColumn,
                        styleIndex,
                        resolvedFormulaTexts, formulas, richText, textRuns);
                    if (writeCell != null) result.Add(writeCell);
                    sequentialColumn = cellColumn + 1;
                }
            }

            if (sourceSheet != null) {
                foreach (XlsbCell sourceCell in sourceSheet.Cells) {
                    if (sourceCell.Kind == XlsbCellValueKind.Blank
                        && !visitedSourceCells.Contains((sourceCell.Row, sourceCell.Column))) {
                        result.Add(XlsbWriteCell.PreserveSource(sourceCell));
                    }
                }
            }

            result.Sort(static (left, right) => {
                int row = left.Row.CompareTo(right.Row);
                return row != 0 ? row : left.Column.CompareTo(right.Column);
            });
            return ApplyArrayFormulaGroups(sheet, sheetData, result, formulas).AsReadOnly();
        }

        // Worksheet parts that a feature writer encodes or that the rewrite validates separately: the DrawingML
        // drawing graph (XlsbWorksheetDrawingWriter, XlsbWorksheetDrawingPlan).
        private static void ThrowIfUnsupportedNewWorksheetContent(ExcelSheet sheet) {
            string? unsupportedPart = XlsbWorksheetContent.FindUnsupportedPart(sheet, "relationship-backed content");
            if (unsupportedPart != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support {unsupportedPart} on worksheet '{sheet.Name}'.");
            }

            string? unsupported = XlsbWorksheetContent.FindUnsupported(sheet);
            if (unsupported != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support worksheet metadata '{unsupported}' on worksheet '{sheet.Name}'.");
            }
        }

        private static void ThrowIfUnsupportedWorksheetMutation(ExcelSheet sheet, XlsbSheet sourceSheet) {
            // Comments and their VML drawing are compared with the source by XlsbWorksheetCommentRewrite.
            string? unsupportedPart = XlsbWorksheetContent.FindUnsupportedPart(sheet, "relationship-backed worksheet content");
            if (unsupportedPart != null) {
                throw new NotSupportedException($"Native XLSB rewriting currently cannot modify {unsupportedPart} on worksheet '{sheet.Name}'; save as .xlsx to retain that change.");
            }
            // Edited tables are rewritten in place; adding or removing a table needs new package parts.
            if (!XlsbWorksheetTableProjector.TryPair(sheet, sourceSheet, out _)) {
                throw new NotSupportedException($"Native XLSB rewriting cannot add or remove tables on worksheet '{sheet.Name}'; save as .xlsx to retain that change.");
            }

            XlsbWorksheetGeometryProjector.ValidateUnchanged(sheet, sourceSheet);
            XlsbWorksheetPropertiesProjector.ValidateUnchanged(sheet, sourceSheet.Properties);
            XlsbWorksheetProtectionProjector.ValidateUnchanged(sheet, sourceSheet.Protection);
            XlsbWorksheetPrintSettingsProjector.ValidateUnchanged(
                sheet,
                sourceSheet.PrintOptions,
                sourceSheet.PageMargins,
                sourceSheet.PageSetup,
                sourceSheet.HeaderFooter);
        }

        private static XlsbWriteCell? ConvertCell(
            ExcelDocument document,
            ExcelSheet sheet,
            CellMarkup cell,
            XlsbCell? sourceCell,
            int row,
            int column,
            uint styleIndex,
            IReadOnlyDictionary<string, string> resolvedFormulaTexts, IXlsbFormulaEncodingContext? formulas,
            XlsbRichTextFontTable? richText, XlsbCellTextRuns textRuns) {
            if (cell.Formula.Exists) {
                return ConvertFormulaCell(document, sheet, cell, sourceCell, row, column, styleIndex, resolvedFormulaTexts, formulas, textRuns);
            }

            string? dataType = cell.DataType;
            string rawValue = cell.Value.InnerText;
            if (dataType == "b") {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Boolean, rawValue == "1" || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase));
            }

            if (dataType == "e") {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Error, GetErrorCode(rawValue, row, column));
            }

            if (dataType == "d") {
                if (!ExcelIsoDateText.TryGetSerial(rawValue, document.DateSystem, DateTimeStyles.RoundtripKind, out double serial)) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot encode date value '{rawValue}' at {ToAddress(row, column)}.");
                }

                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Number, serial);
            }

            if ((dataType == "n" || dataType == null)
                && OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Number, number);
            }

            if (dataType == "s"
                || dataType == "inlineStr"
                || dataType == "str"
                || !string.IsNullOrEmpty(rawValue)) {
                string text = textRuns.ReadText(sheet, cell.Node, dataType, rawValue);
                EnsureCellTextLength(text, row, column);
                if (richText == null && textRuns.HasFormattedRuns(cell.Node)) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot add rich text runs at {sheet.Name}!{ToAddress(row, column)}; save as .xlsx to retain that change.");
                }
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Text, text) {
                    TextRuns = richText?.CreateRuns(cell.Node, styleIndex, text)
                };
            }

            // Preserve native source cells and intentional styled blanks, including
            // tabular missing cells carrying the explicit default format.
            return sourceCell != null || cell.StyleIndex.HasValue
                ? new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Blank, null)
                : null;
        }

        private static XlsbWriteCell ConvertFormulaCell(
            ExcelDocument document,
            ExcelSheet sheet,
            CellMarkup cell,
            XlsbCell? sourceCell,
            int row,
            int column,
            uint styleIndex,
            IReadOnlyDictionary<string, string> resolvedFormulaTexts, IXlsbFormulaEncodingContext? formulas,
            XlsbCellTextRuns textRuns) {
            string formulaText = GetResolvedFormulaText(cell, resolvedFormulaTexts);
            byte[] formulaPayload;
            if (sourceCell?.FormulaBytes != null
                && !string.IsNullOrWhiteSpace(sourceCell.FormulaText)
                && string.Equals(sourceCell.FormulaText, formulaText, StringComparison.Ordinal)) {
                formulaPayload = sourceCell.FormulaPayloadBytes
                    ?? throw new InvalidDataException($"The preserved XLSB formula at {ToAddress(row, column)} has no complete formula payload.");
                formulaPayload = formulas?.RebindSourceFormula(formulaPayload) ?? formulaPayload;
            } else {
                string? reason;
                bool encoded = formulas == null ? XlsbFormulaEncoder.TryEncode(formulaText, out formulaPayload, out reason)
                    : formulas.TryEncode(formulaText, sheet.Name, out formulaPayload, out reason);
                if (!encoded)
                throw new NotSupportedException($"Native XLSB generation cannot encode formula '{formulaText}' at {ToAddress(row, column)}. {reason}");
            }

            if (XlsbWorkbookMarkup.ReadBoolean(cell.Formula, "ca") == true && (sourceCell?.FormulaBytes == null || sourceCell.FormulaText != formulaText)) { formulaPayload = (byte[])formulaPayload.Clone(); formulaPayload[0] |= 2; }
            if (cell.Value.IsNone && cell.InlineString.IsNone) {
                formulaPayload = (byte[])formulaPayload.Clone(); formulaPayload[0] |= 2;
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaText, string.Empty, formulaPayload);
            }

            string? dataType = cell.DataType;
            string rawValue = cell.Value.InnerText;
            if (dataType == "b") {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaBoolean, rawValue == "1" || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase), formulaPayload);
            }

            if (dataType == "e") {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaError, GetErrorCode(rawValue, row, column), formulaPayload);
            }

            if (dataType == "s" || dataType == "inlineStr" || dataType == "str") {
                string text = textRuns.ReadText(sheet, cell.Node, dataType, rawValue);
                EnsureCellTextLength(text, row, column);
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaText, text, formulaPayload);
            }

            if (dataType == "d"
                && ExcelIsoDateText.TryGetSerial(rawValue, document.DateSystem, DateTimeStyles.RoundtripKind, out double serial)) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaNumber, serial, formulaPayload);
            }

            if (OfficeIMO.Core.Internal.OfficeRoundTripText.TryParseFinite(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                // BrtFmlaNum holds an Xnum, which has no negative zero.
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaNumber, number == 0D ? 0D : number, formulaPayload);
            }

            return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaNumber, 0D, formulaPayload);
        }

        private static string GetResolvedFormulaText(
            CellMarkup cell,
            IReadOnlyDictionary<string, string> resolvedFormulaTexts) {
            string? reference = cell.Reference;
            return reference != null && resolvedFormulaTexts.TryGetValue(reference, out string? formulaText)
                ? formulaText
                : cell.Formula.InnerText;
        }

        private static bool CellMatchesSource(
            ExcelSheet sheet,
            CellMarkup cell,
            XlsbCell sourceCell,
            IReadOnlyDictionary<string, string> resolvedFormulaTexts,
            XlsbCellTextRuns textRuns) {
            string? currentFormula = cell.Formula.IsNone
                ? null
                : GetResolvedFormulaText(cell, resolvedFormulaTexts);
            if (sourceCell.FormulaBytes != null) {
                if (sourceCell.FormulaText == null) {
                    if (currentFormula != null) return false;
                } else if (!string.Equals(sourceCell.FormulaText, currentFormula, StringComparison.Ordinal)) {
                    return false;
                }
            } else if (currentFormula != null) {
                return false;
            }

            string rawValue = cell.Value.InnerText;
            switch (sourceCell.Kind) {
                case XlsbCellValueKind.Blank:
                    return string.IsNullOrEmpty(rawValue)
                        && cell.InlineString.IsNone;
                case XlsbCellValueKind.Number:
                    return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                        && sourceCell.Value is double sourceNumber
                        && number.Equals(sourceNumber);
                case XlsbCellValueKind.Text:
                    // Equal text with added, changed or removed runs is an edit, so the source record is not kept.
                    return string.Equals(textRuns.ReadText(sheet, cell.Node, cell.DataType, rawValue), sourceCell.Value as string ?? string.Empty, StringComparison.Ordinal)
                        && (sourceCell.FormulaBytes != null || textRuns.MatchesSource(cell.Node, sourceCell));
                case XlsbCellValueKind.Boolean:
                    bool value = rawValue == "1" || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase);
                    return sourceCell.Value is bool sourceBoolean && value == sourceBoolean;
                case XlsbCellValueKind.Error:
                    return string.Equals(rawValue, sourceCell.Value as string, StringComparison.Ordinal);
                default:
                    return false;
            }
        }

        private static uint ResolveStyleIndex(
            CellMarkup cell,
            XlsbCell? sourceCell,
            string sheetName,
            int row,
            int column,
            bool allowNewStyles,
            int? availableCellFormatCount) {
            uint current = cell.StyleIndex ?? 0U;
            if (sourceCell != null) {
                if (current != sourceCell.StyleIndex
                    && (!availableCellFormatCount.HasValue || current >= availableCellFormatCount.Value)) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot apply unavailable style index {current} to cell {sheetName}!{ToAddress(row, column)}.");
                }

                return current;
            }

            if (current != 0
                && !allowNewStyles
                && (!availableCellFormatCount.HasValue || current >= availableCellFormatCount.Value)) {
                throw new NotSupportedException($"Native XLSB rewriting cannot add cell {sheetName}!{ToAddress(row, column)} with unavailable style index {current}.");
            }

            return current;
        }

        private static OfficeXmlNode.ChildList WorksheetChildren(OfficeXmlNode parent, string name) =>
            parent.Children(OfficeXmlNamespaces.SpreadsheetML, name);

        private readonly struct CellMarkup {
            internal CellMarkup(OfficeXmlNode node) => Node = node;

            internal OfficeXmlNode Node { get; }
            internal string? Reference => Node.Attribute("r");
            internal string? DataType => XlsbCellTextRuns.ReadCellType(Node);
            internal uint? StyleIndex => XlsbWorkbookMarkup.ReadUInt32(Node, "s");
            internal OfficeXmlNode Formula => XlsbWorksheetMarkup.Child(Node, "f");
            internal OfficeXmlNode Value => XlsbWorksheetMarkup.Child(Node, "v");
            internal OfficeXmlNode InlineString => XlsbWorksheetMarkup.Child(Node, "is");
        }

        private static byte GetErrorCode(string value, int row, int column) {
            if (ExcelErrorCode.TryGetCode(value, out byte code)) return code;
            throw new NotSupportedException($"Native XLSB rewriting cannot encode error value '{value}' at {ToAddress(row, column)}.");
        }

        private static void EnsureCellTextLength(string text, int row, int column) {
            if (text.Length > 32_767) {
                throw new NotSupportedException($"Native XLSB rewriting supports at most 32,767 characters in cell {ToAddress(row, column)}.");
            }
        }

        private static void ParseCellReference(string? reference, uint fallbackRow, int fallbackColumn, out int row, out int column) {
            if (string.IsNullOrWhiteSpace(reference)) {
                row = checked((int)fallbackRow);
                column = fallbackColumn;
                return;
            }

            int index = 0;
            int parsedColumn = 0;
            while (index < reference!.Length && char.IsLetter(reference[index])) {
                parsedColumn = checked(parsedColumn * 26 + (char.ToUpperInvariant(reference[index]) - 'A' + 1));
                index++;
            }

            if (index == 0
                || index == reference.Length
                || !int.TryParse(reference.Substring(index), NumberStyles.None, CultureInfo.InvariantCulture, out int parsedRow)) {
                throw new NotSupportedException($"Native XLSB rewriting cannot parse cell reference '{reference}'.");
            }

            row = parsedRow;
            column = parsedColumn;
        }

        private static string ToAddress(int row, int column) {
            int value = column;
            var name = new StringBuilder();
            while (value > 0) {
                value--;
                name.Insert(0, (char)('A' + value % 26));
                value /= 26;
            }

            return name + row.ToString(CultureInfo.InvariantCulture);
        }
    }
}
