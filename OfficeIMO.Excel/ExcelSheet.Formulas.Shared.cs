using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
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

        private IReadOnlyDictionary<uint, SharedFormulaDefinition> BuildSharedFormulaDefinitions(
            IReadOnlyDictionary<Cell, (int Row, int Column)>? effectiveCoordinates = null,
            IEnumerable<Cell>? cells = null) {
            effectiveCoordinates ??= BuildEffectiveCellCoordinates();
            cells ??= WorksheetRoot.Descendants<Cell>();
            var definitions = new Dictionary<uint, SharedFormulaDefinition>();
            foreach (Cell cell in cells) {
                CellFormula? cellFormula = cell.CellFormula;
                if (cellFormula?.FormulaType?.Value != CellFormulaValues.Shared
                    || cellFormula.SharedIndex?.Value is not uint sharedIndex
                    || string.IsNullOrWhiteSpace(cellFormula.Text)
                    || cellFormula.Text.Length > MaxSupportedFormulaLength
                    || !effectiveCoordinates.TryGetValue(cell, out var coordinate)) {
                    continue;
                }

                if (!definitions.ContainsKey(sharedIndex)) {
                    definitions.Add(
                        sharedIndex,
                        new SharedFormulaDefinition(
                            coordinate.Row,
                            coordinate.Column,
                            cellFormula.Text,
                            cellFormula.Reference?.Value));
                }
            }

            return definitions;
        }

        private string ResolveCellFormulaText(
            Cell cell,
            IReadOnlyDictionary<uint, SharedFormulaDefinition>? sharedFormulaDefinitions = null,
            IReadOnlyDictionary<Cell, (int Row, int Column)>? effectiveCoordinates = null) {
            CellFormula? cellFormula = cell.CellFormula;
            if (cellFormula == null) {
                return string.Empty;
            }

            string formula = cellFormula.Text ?? string.Empty;
            if (cellFormula.FormulaType?.Value != CellFormulaValues.Shared || formula.Length > 0) {
                return formula;
            }

            if (cellFormula.SharedIndex?.Value is not uint sharedIndex) {
                return string.Empty;
            }

            int row;
            int column;
            if (!A1.TryParseCellReferenceFast(cell.CellReference?.Value, out row, out column)) {
                effectiveCoordinates ??= BuildEffectiveCellCoordinates();
                if (!effectiveCoordinates.TryGetValue(cell, out var coordinate)) {
                    return string.Empty;
                }

                row = coordinate.Row;
                column = coordinate.Column;
            }

            sharedFormulaDefinitions ??= BuildSharedFormulaDefinitions(effectiveCoordinates);
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
            IReadOnlyDictionary<Cell, (int Row, int Column)> effectiveCoordinates = BuildEffectiveCellCoordinates();
            IReadOnlyDictionary<uint, SharedFormulaDefinition> sharedFormulaDefinitions =
                BuildSharedFormulaDefinitions(effectiveCoordinates);
            var formulas = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Cell cell in WorksheetRoot.Descendants<Cell>().Where(candidate => candidate.CellFormula != null)) {
                if (effectiveCoordinates.TryGetValue(cell, out var coordinate)) {
                    formulas[BuildCellReference(coordinate.Row, coordinate.Column)] =
                        ResolveCellFormulaText(cell, sharedFormulaDefinitions, effectiveCoordinates);
                }
            }

            return formulas;
        }

        private IReadOnlyList<string> ResolveSharedFormulaTextsForStructuralValidation() {
            List<Cell> sharedCells = WorksheetRoot.Descendants<Cell>()
                .Where(candidate => candidate.CellFormula?.FormulaType?.Value == CellFormulaValues.Shared
                    && candidate.CellFormula.SharedIndex?.Value != null)
                .ToList();
            if (sharedCells.Count == 0) {
                return Array.Empty<string>();
            }

            IReadOnlyDictionary<Cell, (int Row, int Column)> effectiveCoordinates = BuildEffectiveCellCoordinates();
            IReadOnlyDictionary<uint, SharedFormulaDefinition> definitions =
                BuildSharedFormulaDefinitions(effectiveCoordinates);

            var resolved = new List<string>();
            foreach (Cell cell in sharedCells) {
                string text = ResolveCellFormulaText(cell, definitions, effectiveCoordinates);
                if (string.IsNullOrWhiteSpace(text)) {
                    throw new InvalidOperationException(
                        $"Cannot edit rows because shared formula group {cell.CellFormula!.SharedIndex!.Value} cannot be validated safely.");
                }
                resolved.Add(text);
            }
            return resolved;
        }

        private void MaterializeWorkbookSharedFormulasForStructuralEdit() {
            foreach (Sheet sheetElement in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (sheetElement.Id?.Value is not string relationshipId
                    || WorkbookPartRoot.GetPartById(relationshipId) is not DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart) {
                    continue;
                }

                ExcelSheet sheet = ReferenceEquals(worksheetPart, _worksheetPart)
                    ? this
                    : new ExcelSheet(
                        _excelDocument,
                        _spreadSheetDocument,
                        sheetElement,
                        registerSheetWrapper: false);
                sheet.MaterializeSharedFormulasForStructuralEdit();
            }
        }

        private void ValidateWorkbookSharedFormulasForStructuralEdit() {
            foreach (Sheet sheetElement in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (sheetElement.Id?.Value is not string relationshipId
                    || WorkbookPartRoot.GetPartById(relationshipId) is not DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart) {
                    continue;
                }

                ExcelSheet sheet = ReferenceEquals(worksheetPart, _worksheetPart)
                    ? this
                    : new ExcelSheet(
                        _excelDocument,
                        _spreadSheetDocument,
                        sheetElement,
                        registerSheetWrapper: false);
                sheet.ResolveSharedFormulaTextsForStructuralValidation();
            }
        }

        private void MaterializeSharedFormulasForStructuralEdit() {
            List<Cell> sharedCells = WorksheetRoot.Descendants<Cell>()
                .Where(cell => cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Shared
                    && cell.CellFormula.SharedIndex?.Value != null)
                .ToList();
            if (sharedCells.Count == 0) {
                return;
            }

            IReadOnlyDictionary<Cell, (int Row, int Column)> effectiveCoordinates = BuildEffectiveCellCoordinates();
            IReadOnlyDictionary<uint, SharedFormulaDefinition> definitions =
                BuildSharedFormulaDefinitions(effectiveCoordinates);
            var resolved = new List<(CellFormula Formula, string Text)>();
            foreach (Cell cell in sharedCells) {
                CellFormula formula = cell.CellFormula!;
                if (formula.SharedIndex?.Value is not uint sharedIndex) {
                    continue;
                }

                string text = ResolveCellFormulaText(cell, definitions, effectiveCoordinates);
                if (string.IsNullOrWhiteSpace(text)) {
                    throw new InvalidOperationException(
                        $"Cannot edit rows because shared formula group {sharedIndex} cannot be materialized safely.");
                }

                resolved.Add((formula, text));
            }

            foreach ((CellFormula formula, string text) in resolved) {
                formula.Text = text;
                formula.FormulaType = null;
                formula.SharedIndex = null;
                formula.Reference = null;
                formula.CalculateCell = true;
            }
        }

        private IReadOnlyDictionary<Cell, (int Row, int Column)> BuildEffectiveCellCoordinates(
            IEnumerable<Row>? rows = null,
            IEnumerable<Cell>? cells = null) {
            var coordinates = new Dictionary<Cell, (int Row, int Column)>();
            uint previousRow = 0U;
            rows ??= WorksheetRoot.GetFirstChild<SheetData>()?.Elements<Row>()
                ?? Enumerable.Empty<Row>();
            ILookup<OpenXmlElement?, Cell>? cellsByRow =
                cells?.ToLookup(cell => cell.Parent);
            foreach (Row row in rows) {
                uint effectiveRow = GetEffectiveRowIndex(row, previousRow);
                int previousColumn = 0;
                IEnumerable<Cell> rowCells = cellsByRow == null
                    ? row.Elements<Cell>()
                    : cellsByRow[row];
                foreach (Cell cell in rowCells) {
                    int effectiveColumn = previousColumn + 1;
                    if (A1.TryParseCellReferenceFast(
                        cell.CellReference?.Value,
                        out int explicitRow,
                        out int explicitColumn)) {
                        coordinates[cell] = (explicitRow, explicitColumn);
                        effectiveColumn = explicitColumn;
                    } else if (effectiveRow <= (uint)A1.MaxRows
                        && effectiveColumn <= A1.MaxColumns) {
                        coordinates[cell] = (checked((int)effectiveRow), effectiveColumn);
                    }

                    previousColumn = effectiveColumn;
                }

                previousRow = effectiveRow;
            }

            return coordinates;
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
            if (formula.Length == 0 || formula.Length > MaxSupportedFormulaLength
                || (rowOffset == 0 && columnOffset == 0)) {
                return formula;
            }

            return RewriteFormulaReferencesOutsideStrings(formula, segment =>
                ExcelFormulaReferenceRewriter.SharedFormulaReferenceRegex.Replace(segment, match =>
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
            if (ExcelFormulaCapabilities.IsBuiltInFunction(token)
                || _excelDocument.Calculation.TryGetCustomFunction(token, out _)) {
                return true;
            }

            // An unregistered cell-like token followed by a parenthesized reference
            // is Excel's whitespace intersection form. Named functions remain
            // protected above, while both single-cell and range operands translate.
            return !IsParenthesizedReferenceOperand(formula, cursor);
        }

        private static bool IsParenthesizedReferenceOperand(string formula, int openingParenthesis) {
            int cursor = openingParenthesis + 1;
            int depth = 1;
            bool sawReference = false;
            while (cursor < formula.Length) {
                if (char.IsWhiteSpace(formula[cursor]) || formula[cursor] == ',') {
                    cursor++;
                    continue;
                }

                if (formula[cursor] == '(') {
                    depth++;
                    cursor++;
                    continue;
                }

                if (formula[cursor] == ')') {
                    depth--;
                    cursor++;
                    if (depth == 0) {
                        return sawReference;
                    }
                    continue;
                }

                Match reference = ExcelFormulaReferenceRewriter.SharedFormulaReferenceRegex.Match(formula, cursor);
                if (!reference.Success || reference.Index != cursor) {
                    return false;
                }

                sawReference = true;
                cursor += reference.Length;
            }

            return false;
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
            foreach (Match match in ExcelFormulaReferenceRewriter.SharedFormulaReferenceRegex.Matches(formula)) {
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
