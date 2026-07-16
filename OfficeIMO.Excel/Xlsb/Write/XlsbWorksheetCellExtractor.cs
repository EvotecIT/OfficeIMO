using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb.Model;
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

            ThrowIfUnsupportedWorksheetMutation(sheet);
            var sourceCells = sourceSheet.Cells.ToDictionary(cell => (cell.Row, cell.Column));
            var result = new List<XlsbWriteCell>();
            SheetData? sheetData = sheet.WorksheetPart.Worksheet?.GetFirstChild<SheetData>();
            if (sheetData == null) return result;

            foreach (Row row in sheetData.Elements<Row>()) {
                ThrowIfUnsupportedRowMutation(row, sheet.Name);
                uint sequentialRow = row.RowIndex?.Value ?? 0U;
                int sequentialColumn = 1;
                foreach (Cell cell in row.Elements<Cell>()) {
                    ParseCellReference(cell.CellReference?.Value, sequentialRow, sequentialColumn, out int cellRow, out int cellColumn);
                    if (cellRow <= 0 || cellRow > 1_048_576 || cellColumn <= 0 || cellColumn > 16_384) {
                        throw new NotSupportedException($"Native XLSB rewriting cannot encode invalid cell reference '{cell.CellReference?.Value}'.");
                    }

                    sourceCells.TryGetValue((cellRow, cellColumn), out XlsbCell? sourceCell);
                    uint styleIndex = ResolveStyleIndex(cell, sourceCell, sheet.Name, cellRow, cellColumn);
                    if (sourceCell != null && CellMatchesSource(sheet, cell, sourceCell)) {
                        result.Add(XlsbWriteCell.PreserveSource(sourceCell));
                        sequentialColumn = cellColumn + 1;
                        continue;
                    }

                    XlsbWriteCell? writeCell = ConvertCell(document, sheet, cell, sourceCell, cellRow, cellColumn, styleIndex);
                    if (writeCell != null) result.Add(writeCell);
                    sequentialColumn = cellColumn + 1;
                }
            }

            result.Sort(static (left, right) => {
                int row = left.Row.CompareTo(right.Row);
                return row != 0 ? row : left.Column.CompareTo(right.Column);
            });
            return result.AsReadOnly();
        }

        private static void ThrowIfUnsupportedWorksheetMutation(ExcelSheet sheet) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            OpenXmlElement? unsupportedChild = worksheet.ChildElements.FirstOrDefault(element => element is not SheetData);
            if (unsupportedChild != null
                || sheet.WorksheetPart.Parts.Any()
                || sheet.WorksheetPart.ExternalRelationships.Any()
                || sheet.WorksheetPart.HyperlinkRelationships.Any()) {
                string detail = unsupportedChild?.LocalName ?? "relationship-backed worksheet content";
                throw new NotSupportedException($"Native XLSB rewriting currently accepts cell-value edits only. Worksheet '{sheet.Name}' contains modified {detail}; save as .xlsx to retain that change.");
            }
        }

        private static void ThrowIfUnsupportedRowMutation(Row row, string sheetName) {
            if (row.CustomFormat?.Value == true
                || row.CustomHeight?.Value == true
                || row.Hidden?.Value == true
                || (row.OutlineLevel?.Value ?? 0) != 0
                || row.Collapsed?.Value == true
                || row.StyleIndex != null) {
                throw new NotSupportedException($"Native XLSB rewriting currently accepts cell-value edits only. Worksheet '{sheetName}' contains modified row formatting or outline metadata.");
            }
        }

        private static XlsbWriteCell? ConvertCell(
            ExcelDocument document,
            ExcelSheet sheet,
            Cell cell,
            XlsbCell? sourceCell,
            int row,
            int column,
            uint styleIndex) {
            if (cell.CellFormula != null) {
                return ConvertFormulaCell(document, sheet, cell, sourceCell, row, column, styleIndex);
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
            uint styleIndex) {
            string formulaText = cell.CellFormula?.Text ?? string.Empty;
            if (sourceCell?.FormulaBytes == null
                || string.IsNullOrWhiteSpace(sourceCell.FormulaText)
                || !string.Equals(sourceCell.FormulaText, formulaText, StringComparison.Ordinal)) {
                throw new NotSupportedException($"Native XLSB rewriting currently preserves existing formula token streams but does not encode changed formula '{formulaText}' at {ToAddress(row, column)}. Save as .xlsx or keep the source formula unchanged.");
            }

            CellValues? dataType = cell.DataType?.Value;
            string rawValue = cell.CellValue?.InnerText ?? string.Empty;
            if (dataType == CellValues.Boolean) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaBoolean, rawValue == "1" || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase), sourceCell.FormulaPayloadBytes);
            }

            if (dataType == CellValues.Error) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaError, GetErrorCode(rawValue, row, column), sourceCell.FormulaPayloadBytes);
            }

            if (dataType == CellValues.SharedString || dataType == CellValues.InlineString || dataType == CellValues.String) {
                string text = sheet.GetCellText(cell);
                EnsureCellTextLength(text, row, column);
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaText, text, sourceCell.FormulaPayloadBytes);
            }

            if (dataType == CellValues.Date
                && DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime date)) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaNumber, ExcelDateSystemConverter.ToSerial(date, document.DateSystem), sourceCell.FormulaPayloadBytes);
            }

            if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                return new XlsbWriteCell(row, column, styleIndex, XlsbWriteCellKind.FormulaNumber, number, sourceCell.FormulaPayloadBytes);
            }

            throw new NotSupportedException($"Native XLSB rewriting requires a cached result for formula cell {ToAddress(row, column)}.");
        }

        private static bool CellMatchesSource(ExcelSheet sheet, Cell cell, XlsbCell sourceCell) {
            string? currentFormula = cell.CellFormula?.Text;
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

        private static uint ResolveStyleIndex(Cell cell, XlsbCell? sourceCell, string sheetName, int row, int column) {
            uint current = cell.StyleIndex?.Value ?? 0U;
            if (current != 0) {
                throw new NotSupportedException($"Native XLSB rewriting currently accepts cell-value edits only. Cell {sheetName}!{ToAddress(row, column)} has modified style index {current}.");
            }

            return sourceCell?.StyleIndex ?? 0U;
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
