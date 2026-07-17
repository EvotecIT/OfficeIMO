using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static readonly Regex SharedFormulaReferenceRegex = new Regex(
            @"(?<![A-Za-z0-9_\.])(?<qualifier>(?:'(?:[^']|'')+'|\[[^\]]+\][A-Za-z0-9_\. ]+|[A-Za-z_][A-Za-z0-9_\. ]*:[A-Za-z_][A-Za-z0-9_\. ]*|[A-Za-z_][A-Za-z0-9_\. ]*)!)?(?:(?<cellStartColumnAbsolute>\$?)(?<cellStartColumn>[A-Za-z]{1,3})(?<cellStartRowAbsolute>\$?)(?<cellStartRow>\d{1,7})(?::(?<cellEndColumnAbsolute>\$?)(?<cellEndColumn>[A-Za-z]{1,3})(?<cellEndRowAbsolute>\$?)(?<cellEndRow>\d{1,7}))?(?<cellSpill>#)?|(?<wholeStartColumnAbsolute>\$?)(?<wholeStartColumn>[A-Za-z]{1,3}):(?<wholeEndColumnAbsolute>\$?)(?<wholeEndColumn>[A-Za-z]{1,3})|(?<wholeStartRowAbsolute>\$?)(?<wholeStartRow>\d{1,7}):(?<wholeEndRowAbsolute>\$?)(?<wholeEndRow>\d{1,7}))(?![A-Za-z0-9_\.]|\()",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

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

        private IReadOnlyDictionary<uint, SharedFormulaDefinition> GetFormulaEvaluationSharedDefinitions(
            ExcelSheet sheet) {
            Dictionary<string, IReadOnlyDictionary<uint, SharedFormulaDefinition>>? definitionsBySheet =
                _formulaEvaluationSharedDefinitionsBySheet;
            if (definitionsBySheet == null) {
                return sheet.BuildSharedFormulaDefinitions();
            }

            if (!definitionsBySheet.TryGetValue(sheet.Name, out IReadOnlyDictionary<uint, SharedFormulaDefinition>? definitions)) {
                definitions = sheet.BuildSharedFormulaDefinitions();
                definitionsBySheet.Add(sheet.Name, definitions);
            }

            return definitions;
        }

        internal IReadOnlyDictionary<string, string> BuildResolvedFormulaTextMap() {
            IReadOnlyDictionary<uint, SharedFormulaDefinition> sharedFormulaDefinitions = BuildSharedFormulaDefinitions();
            var formulas = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Cell cell in WorksheetRoot.Descendants<Cell>().Where(candidate => candidate.CellFormula != null)) {
                string? cellReference = cell.CellReference?.Value;
                if (!string.IsNullOrWhiteSpace(cellReference)) {
                    formulas[cellReference!] = ResolveCellFormulaText(cell, sharedFormulaDefinitions);
                }
            }

            return formulas;
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

        private string TranslateSharedFormula(string formula, int rowOffset, int columnOffset) {
            if (formula.Length == 0 || (rowOffset == 0 && columnOffset == 0)) {
                return formula;
            }

            return RewriteFormulaReferencesOutsideStrings(formula, segment =>
                SharedFormulaReferenceRegex.Replace(segment, match =>
                    IsInsideFormulaStructuredReference(segment, match.Index)
                        || IsSharedFormulaFunctionToken(segment, match)
                        ? match.Value
                        : TranslateSharedFormulaReference(match, rowOffset, columnOffset)));
        }

        private bool IsSharedFormulaFunctionToken(string formula, Match match) {
            if (match.Groups["qualifier"].Success
                || match.Groups["cellEndColumn"].Success
                || match.Groups["cellSpill"].Success
                || match.Groups["cellStartColumnAbsolute"].Value.Length > 0
                || match.Groups["cellStartRowAbsolute"].Value.Length > 0) {
                return false;
            }

            int cursor = match.Index + match.Length;
            int whitespaceStart = cursor;
            while (cursor < formula.Length && char.IsWhiteSpace(formula[cursor])) {
                cursor++;
            }

            if (cursor == whitespaceStart || cursor >= formula.Length || formula[cursor] != '(') {
                return false;
            }

            string token = match.Groups["cellStartColumn"].Value + match.Groups["cellStartRow"].Value;
            return ExcelFormulaCapabilities.IsBuiltInFunction(token)
                || _excelDocument.Calculation.TryGetCustomFunction(token, out _);
        }

        private static string TranslateSharedFormulaReference(Match match, int rowOffset, int columnOffset) {
            string qualifier = match.Groups["qualifier"].Value;
            if (match.Groups["cellStartColumn"].Success) {
                string? start = TranslateSharedFormulaCell(
                    match.Groups["cellStartColumn"].Value,
                    match.Groups["cellStartColumnAbsolute"].Value,
                    match.Groups["cellStartRow"].Value,
                    match.Groups["cellStartRowAbsolute"].Value,
                    rowOffset,
                    columnOffset);
                if (start == null) {
                    return match.Value;
                }

                string reference = start;
                if (match.Groups["cellEndColumn"].Success) {
                    string? end = TranslateSharedFormulaCell(
                        match.Groups["cellEndColumn"].Value,
                        match.Groups["cellEndColumnAbsolute"].Value,
                        match.Groups["cellEndRow"].Value,
                        match.Groups["cellEndRowAbsolute"].Value,
                        rowOffset,
                        columnOffset);
                    if (end == null) {
                        return match.Value;
                    }

                    reference += ":" + end;
                }

                return qualifier + reference + match.Groups["cellSpill"].Value;
            }

            if (match.Groups["wholeStartColumn"].Success) {
                string? start = TranslateSharedFormulaColumn(
                    match.Groups["wholeStartColumn"].Value,
                    match.Groups["wholeStartColumnAbsolute"].Value,
                    columnOffset);
                string? end = TranslateSharedFormulaColumn(
                    match.Groups["wholeEndColumn"].Value,
                    match.Groups["wholeEndColumnAbsolute"].Value,
                    columnOffset);
                return start == null || end == null
                    ? match.Value
                    : qualifier + start + ":" + end;
            }

            if (match.Groups["wholeStartRow"].Success) {
                string? start = TranslateSharedFormulaRow(
                    match.Groups["wholeStartRow"].Value,
                    match.Groups["wholeStartRowAbsolute"].Value,
                    rowOffset);
                string? end = TranslateSharedFormulaRow(
                    match.Groups["wholeEndRow"].Value,
                    match.Groups["wholeEndRowAbsolute"].Value,
                    rowOffset);
                return start == null || end == null
                    ? match.Value
                    : qualifier + start + ":" + end;
            }

            return match.Value;
        }

        private static string? TranslateSharedFormulaCell(
            string columnText,
            string columnAbsolute,
            string rowText,
            string rowAbsolute,
            int rowOffset,
            int columnOffset) {
            if (!int.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out int sourceRow)
                || !TryParseSharedFormulaColumn(columnText, out int sourceColumn)
                || sourceRow <= 0
                || sourceRow > A1.MaxRows
                || sourceColumn <= 0
                || sourceColumn > A1.MaxColumns) {
                return null;
            }

            int targetRow = rowAbsolute.Length > 0 ? sourceRow : sourceRow + rowOffset;
            int targetColumn = columnAbsolute.Length > 0 ? sourceColumn : sourceColumn + columnOffset;
            if (targetRow <= 0 || targetRow > A1.MaxRows || targetColumn <= 0 || targetColumn > A1.MaxColumns) {
                return "#REF!";
            }

            return columnAbsolute
                + A1.ColumnIndexToLetters(targetColumn)
                + rowAbsolute
                + targetRow.ToString(CultureInfo.InvariantCulture);
        }

        private static string? TranslateSharedFormulaColumn(string columnText, string columnAbsolute, int columnOffset) {
            if (!TryParseSharedFormulaColumn(columnText, out int sourceColumn)
                || sourceColumn <= 0
                || sourceColumn > A1.MaxColumns) {
                return null;
            }

            int targetColumn = columnAbsolute.Length > 0 ? sourceColumn : sourceColumn + columnOffset;
            return targetColumn <= 0 || targetColumn > A1.MaxColumns
                ? "#REF!"
                : columnAbsolute + A1.ColumnIndexToLetters(targetColumn);
        }

        private static string? TranslateSharedFormulaRow(string rowText, string rowAbsolute, int rowOffset) {
            if (!int.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out int sourceRow)
                || sourceRow <= 0
                || sourceRow > A1.MaxRows) {
                return null;
            }

            int targetRow = rowAbsolute.Length > 0 ? sourceRow : sourceRow + rowOffset;
            return targetRow <= 0 || targetRow > A1.MaxRows
                ? "#REF!"
                : rowAbsolute + targetRow.ToString(CultureInfo.InvariantCulture);
        }

        private static bool TryParseSharedFormulaColumn(string columnText, out int column) {
            column = A1.ParseColumnIndexFromCellReferenceWithKnownRowFast(columnText + "1");
            return column > 0;
        }

        private static string MaskFormulaNonLocalReferenceSegments(string formula) {
            char[] masked = formula.ToCharArray();
            foreach (Match match in SharedFormulaReferenceRegex.Matches(formula)) {
                string qualifier = match.Groups["qualifier"].Value;
                if (qualifier.IndexOf('[') < 0 && qualifier.IndexOf(']') < 0 && qualifier.IndexOf(':') < 0) {
                    continue;
                }

                for (int index = match.Index; index < match.Index + match.Length; index++) {
                    masked[index] = ' ';
                }
            }

            return new string(masked);
        }
    }
}
