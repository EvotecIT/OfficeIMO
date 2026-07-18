using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;
using System.Globalization;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Creates the bounded cell-only mutation plan supported by the first XLSB rewriter.</summary>
    internal static class XlsbWorksheetCellExtractor {
        internal static IReadOnlyList<XlsbWriteCell> Extract(
            ExcelDocument document,
            ExcelSheet sheet,
            XlsbWorksheet sourceSheet) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (sourceSheet == null) throw new ArgumentNullException(nameof(sourceSheet));

            ThrowIfUnsupportedWorksheetMutation(sheet, sourceSheet);
            return ExtractCore(document, sheet, sourceSheet);
        }

        internal static IReadOnlyList<XlsbWriteCell> ExtractNew(
            ExcelDocument document,
            ExcelSheet sheet) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));

            ThrowIfUnsupportedNewWorksheetContent(sheet);
            return ExtractCore(document, sheet, sourceSheet: null);
        }

        private static IReadOnlyList<XlsbWriteCell> ExtractCore(
            ExcelDocument document,
            ExcelSheet sheet,
            XlsbWorksheet? sourceSheet) {
            var sourceCells = sourceSheet?.Cells.ToDictionary(cell => (cell.Row, cell.Column))
                ?? new Dictionary<(int Row, int Column), XlsbCell>();
            var visitedSourceCells = new HashSet<(int Row, int Column)>();
            var result = new List<XlsbWriteCell>();
            IReadOnlyDictionary<string, string> resolvedFormulaTexts = sheet.BuildResolvedFormulaTextMap();
            SheetData? sheetData = sheet.WorksheetPart.Worksheet?.GetFirstChild<SheetData>();
            if (sheetData == null) return result;

            foreach (Row row in sheetData.Elements<Row>()) {
                uint sequentialRow = row.RowIndex?.Value ?? 0U;
                int sequentialColumn = 1;
                foreach (Cell cell in row.Elements<Cell>()) {
                    ParseCellReference(cell.CellReference?.Value, sequentialRow, sequentialColumn, out int cellRow, out int cellColumn);
                    if (cellRow <= 0 || cellRow > 1_048_576 || cellColumn <= 0 || cellColumn > 16_384) {
                        throw new NotSupportedException($"Native XLSB rewriting cannot encode invalid cell reference '{cell.CellReference?.Value}'.");
                    }

                    sourceCells.TryGetValue((cellRow, cellColumn), out XlsbCell? sourceCell);
                    if (sourceCell != null) visitedSourceCells.Add((cellRow, cellColumn));
                    uint styleIndex = ResolveStyleIndex(
                        cell,
                        sourceCell,
                        sheet.Name,
                        cellRow,
                        cellColumn,
                        allowNewStyles: sourceSheet == null);
                    if (sourceCell != null && CellMatchesSource(sheet, cell, sourceCell, resolvedFormulaTexts)) {
                        result.Add(XlsbWriteCell.PreserveSource(sourceCell));
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
                        resolvedFormulaTexts);
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
            return result.AsReadOnly();
        }

        private static void ThrowIfUnsupportedNewWorksheetContent(ExcelSheet sheet) {
            if (sheet.WorksheetPart.Parts.Any()
                || sheet.WorksheetPart.ExternalRelationships.Any()) {
                throw new NotSupportedException($"Native XLSB generation does not yet support relationship-backed content on worksheet '{sheet.Name}'.");
            }

            OpenXmlElement? unsupported = sheet.WorksheetPart.Worksheet?.ChildElements
                .FirstOrDefault(element => element is not SheetProperties
                    && element is not SheetDimension
                    && element is not SheetViews
                    && element is not SheetFormatProperties
                    && element is not Columns
                    && element is not SheetData
                    && element is not SheetProtection
                    && element is not AutoFilter
                    && element is not MergeCells
                    && element is not Hyperlinks
                    && element is not PrintOptions
                    && element is not PageMargins
                    && element is not PageSetup
                    && element is not HeaderFooter);
            if (unsupported != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support worksheet metadata '{unsupported.LocalName}' on worksheet '{sheet.Name}'.");
            }
        }

        private static void ThrowIfUnsupportedWorksheetMutation(ExcelSheet sheet, XlsbWorksheet sourceSheet) {
            if (sheet.WorksheetPart.Parts.Any()
                || sheet.WorksheetPart.ExternalRelationships.Any()) {
                throw new NotSupportedException($"Native XLSB rewriting currently cannot modify relationship-backed worksheet content on worksheet '{sheet.Name}'; save as .xlsx to retain that change.");
            }

            XlsbWorksheetGeometryProjector.ValidateUnchanged(sheet, sourceSheet);
            XlsbWorksheetPropertiesProjector.ValidateUnchanged(sheet, sourceSheet.Properties);
            XlsbWorksheetProtectionProjector.ValidateUnchanged(sheet, sourceSheet.Protection);
            XlsbWorksheetAutoFilterProjector.ValidateUnchanged(sheet, sourceSheet.AutoFilter);
            XlsbWorksheetPrintSettingsProjector.ValidateUnchanged(
                sheet,
                sourceSheet.PrintOptions,
                sourceSheet.PageMargins,
                sourceSheet.PageSetup,
                sourceSheet.HeaderFooter);
            XlsbWorksheetHyperlinkProjector.ValidateUnchanged(sheet, sourceSheet);
        }

        private static XlsbWriteCell? ConvertCell(
            ExcelDocument document,
            ExcelSheet sheet,
            Cell cell,
            XlsbCell? sourceCell,
            int row,
            int column,
            uint styleIndex,
            IReadOnlyDictionary<string, string> resolvedFormulaTexts) {
            if (cell.CellFormula != null) {
                return ConvertFormulaCell(document, sheet, cell, sourceCell, row, column, styleIndex, resolvedFormulaTexts);
            }

            if (sourceCell?.FormulaBytes != null) {
                throw new NotSupportedException($"Native XLSB rewriting cannot remove or replace the preserved formula at {ToAddress(row, column)}. Save as .xlsx to change formula structure.");
            }

            CellValues? dataType = cell.DataType?.Value;
            string rawValue = cell.CellValue?.InnerText ?? string.Empty;
            if (dataType == CellValues.Boolean) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Boolean, rawValue == "1" || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase));
            }

            if (dataType == CellValues.Error) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Error, GetErrorCode(rawValue, row, column));
            }

            if (dataType == CellValues.Date) {
                if (!DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime date)) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot encode date value '{rawValue}' at {ToAddress(row, column)}.");
                }

                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Number, ExcelDateSystemConverter.ToSerial(date, document.DateSystem));
            }

            if ((dataType == CellValues.Number || dataType == null)
                && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Number, number);
            }

            if (dataType == CellValues.SharedString
                || dataType == CellValues.InlineString
                || dataType == CellValues.String
                || !string.IsNullOrEmpty(rawValue)) {
                string text = sheet.GetCellText(cell);
                EnsureCellTextLength(text, row, column);
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Text, text);
            }

            return sourceCell != null || styleIndex != 0
                ? new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.Blank, null)
                : null;
        }

        private static XlsbWriteCell ConvertFormulaCell(
            ExcelDocument document,
            ExcelSheet sheet,
            Cell cell,
            XlsbCell? sourceCell,
            int row,
            int column,
            uint styleIndex,
            IReadOnlyDictionary<string, string> resolvedFormulaTexts) {
            string formulaText = GetResolvedFormulaText(cell, resolvedFormulaTexts);
            byte[] formulaPayload;
            if (sourceCell?.FormulaBytes != null
                && !string.IsNullOrWhiteSpace(sourceCell.FormulaText)
                && string.Equals(sourceCell.FormulaText, formulaText, StringComparison.Ordinal)) {
                formulaPayload = sourceCell.FormulaPayloadBytes
                    ?? throw new InvalidDataException($"The preserved XLSB formula at {ToAddress(row, column)} has no complete formula payload.");
            } else if (!XlsbFormulaEncoder.TryEncode(formulaText, out formulaPayload, out string? reason)) {
                throw new NotSupportedException($"Native XLSB generation cannot encode formula '{formulaText}' at {ToAddress(row, column)}. {reason}");
            }

            CellValues? dataType = cell.DataType?.Value;
            string rawValue = cell.CellValue?.InnerText ?? string.Empty;
            if (dataType == CellValues.Boolean) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaBoolean, rawValue == "1" || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase), formulaPayload);
            }

            if (dataType == CellValues.Error) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaError, GetErrorCode(rawValue, row, column), formulaPayload);
            }

            if (dataType == CellValues.SharedString || dataType == CellValues.InlineString || dataType == CellValues.String) {
                string text = sheet.GetCellText(cell);
                EnsureCellTextLength(text, row, column);
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaText, text, formulaPayload);
            }

            if (dataType == CellValues.Date
                && DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime date)) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaNumber, ExcelDateSystemConverter.ToSerial(date, document.DateSystem), formulaPayload);
            }

            if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaNumber, number, formulaPayload);
            }

            return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaNumber, 0D, formulaPayload);
        }

        private static string GetResolvedFormulaText(
            Cell cell,
            IReadOnlyDictionary<string, string> resolvedFormulaTexts) {
            string? reference = cell.CellReference?.Value;
            return reference != null && resolvedFormulaTexts.TryGetValue(reference, out string? formulaText)
                ? formulaText
                : cell.CellFormula?.Text ?? string.Empty;
        }

        private static bool CellMatchesSource(
            ExcelSheet sheet,
            Cell cell,
            XlsbCell sourceCell,
            IReadOnlyDictionary<string, string> resolvedFormulaTexts) {
            string? currentFormula = cell.CellFormula == null
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

            string rawValue = cell.CellValue?.InnerText ?? string.Empty;
            switch (sourceCell.Kind) {
                case XlsbCellValueKind.Blank:
                    return string.IsNullOrEmpty(rawValue)
                        && cell.InlineString == null;
                case XlsbCellValueKind.Number:
                    return double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                        && sourceCell.Value is double sourceNumber
                        && number.Equals(sourceNumber);
                case XlsbCellValueKind.Text:
                    return string.Equals(sheet.GetCellText(cell), sourceCell.Value as string ?? string.Empty, StringComparison.Ordinal);
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
            Cell cell,
            XlsbCell? sourceCell,
            string sheetName,
            int row,
            int column,
            bool allowNewStyles) {
            uint current = cell.StyleIndex?.Value ?? 0U;
            if (sourceCell != null) {
                if (current != sourceCell.StyleIndex) {
                    throw new NotSupportedException($"Native XLSB rewriting currently accepts cell-value edits only. Cell {sheetName}!{ToAddress(row, column)} changed style index from {sourceCell.StyleIndex} to {current}.");
                }

                return current;
            }

            if (current != 0 && !allowNewStyles) {
                throw new NotSupportedException($"Native XLSB rewriting currently cannot add a styled cell. Cell {sheetName}!{ToAddress(row, column)} uses style index {current}.");
            }

            return current;
        }

        private static byte GetErrorCode(string value, int row, int column) {
            if (LegacyXlsErrorValue.TryGetCode(value, out byte code)) return code;
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
