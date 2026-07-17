using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private sealed class SharedFormulaDefinition {
            internal SharedFormulaDefinition(int row, int column, string formula, string? reference) {
                Row = row;
                Column = column;
                Formula = formula;
                Reference = reference;
            }

            internal int Row { get; }
            internal int Column { get; }
            internal string Formula { get; }
            internal string? Reference { get; }
        }

        private IReadOnlyDictionary<uint, SharedFormulaDefinition> BuildSharedFormulaDefinitions() {
            var definitions = new Dictionary<uint, SharedFormulaDefinition>();
            foreach (Cell cell in WorksheetRoot.Descendants<Cell>()) {
                CellFormula? cellFormula = cell.CellFormula;
                if (cellFormula?.FormulaType?.Value != CellFormulaValues.Shared
                    || cellFormula.SharedIndex?.Value is not uint sharedIndex
                    || string.IsNullOrWhiteSpace(cellFormula.Text)
                    || !A1.TryParseCellReferenceFast(cell.CellReference?.Value, out int row, out int column)) {
                    continue;
                }

                if (!definitions.ContainsKey(sharedIndex)) {
                    definitions.Add(
                        sharedIndex,
                        new SharedFormulaDefinition(row, column, cellFormula.Text, cellFormula.Reference?.Value));
                }
            }

            return definitions;
        }

        private string ResolveCellFormulaText(
            Cell cell,
            IReadOnlyDictionary<uint, SharedFormulaDefinition>? sharedFormulaDefinitions = null) {
            CellFormula? cellFormula = cell.CellFormula;
            if (cellFormula == null) {
                return string.Empty;
            }

            string formula = cellFormula.Text ?? string.Empty;
            if (cellFormula.FormulaType?.Value != CellFormulaValues.Shared || formula.Length > 0) {
                return formula;
            }

            if (cellFormula.SharedIndex?.Value is not uint sharedIndex
                || !A1.TryParseCellReferenceFast(cell.CellReference?.Value, out int row, out int column)) {
                return string.Empty;
            }

            sharedFormulaDefinitions ??= BuildSharedFormulaDefinitions();
            if (!sharedFormulaDefinitions.TryGetValue(sharedIndex, out SharedFormulaDefinition? definition)
                || !ContainsSharedFormulaCell(definition, row, column)) {
                return string.Empty;
            }

            return TranslateSharedFormula(
                definition.Formula,
                row - definition.Row,
                column - definition.Column);
        }

        private static bool ContainsSharedFormulaCell(SharedFormulaDefinition definition, int row, int column) {
            if (string.IsNullOrWhiteSpace(definition.Reference)) {
                return true;
            }

            return A1.TryParseRange(
                    definition.Reference!.Replace("$", string.Empty),
                    out int firstRow,
                    out int firstColumn,
                    out int lastRow,
                    out int lastColumn)
                && row >= firstRow
                && row <= lastRow
                && column >= firstColumn
                && column <= lastColumn;
        }

        private static string TranslateSharedFormula(string formula, int rowOffset, int columnOffset) {
            if (formula.Length == 0 || (rowOffset == 0 && columnOffset == 0)) {
                return formula;
            }

            return RewriteFormulaReferencesOutsideStrings(formula, segment =>
                ReplaceFormulaReferences(segment, match => {
                    if (IsInsideFormulaStructuredReference(segment, match.Index)
                        || !int.TryParse(match.Groups["row"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int sourceRow)) {
                        return match.Value;
                    }

                    int sourceColumn = A1.ParseColumnIndexFromCellReferenceWithKnownRowFast(
                        match.Groups["col"].Value + sourceRow.ToString(CultureInfo.InvariantCulture));
                    if (sourceRow <= 0 || sourceRow > A1.MaxRows || sourceColumn <= 0 || sourceColumn > A1.MaxColumns) {
                        return match.Value;
                    }

                    int targetRow = match.Groups["rowAbs"].Value.Length > 0
                        ? sourceRow
                        : sourceRow + rowOffset;
                    int targetColumn = match.Groups["colAbs"].Value.Length > 0
                        ? sourceColumn
                        : sourceColumn + columnOffset;
                    if (targetRow <= 0 || targetRow > A1.MaxRows || targetColumn <= 0 || targetColumn > A1.MaxColumns) {
                        return "#REF!";
                    }

                    string sheetQualifier = match.Groups["sheet"].Value;
                    return sheetQualifier
                        + (sheetQualifier.Length > 0 ? "!" : string.Empty)
                        + match.Groups["colAbs"].Value
                        + A1.ColumnIndexToLetters(targetColumn)
                        + match.Groups["rowAbs"].Value
                        + targetRow.ToString(CultureInfo.InvariantCulture);
                }));
        }
    }
}
