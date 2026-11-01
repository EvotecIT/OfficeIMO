namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private int _indexFormulaReferenceDepth;

        private bool TryEvaluateIndexValue(string args, out FormulaArgumentValue result) {
            result = default;
            if (!TrySelectIndexSource(args, out FormulaLookupSource selected) || selected.Count != 1) return false;
            result = NormalizeLookupResult(selected[0]);
            return !result.IsUnresolvedFormula;
        }

        private bool TrySelectIndexSource(string args, out FormulaLookupSource selected) {
            selected = null!;
            var tokens = SplitFormulaArguments(args, preserveEmpty: true);
            if (tokens.Count < 2 || tokens.Count > 4) return false;
            int? currentRow = _formulaEvaluationCellReference != null
                && TryParseCellReference(_formulaEvaluationCellReference, out int evaluationRow, out _) ? evaluationRow : null;
            if (TryResolveIndexRange(args, currentRow, out ExcelSheet sheet, out int r1, out int c1,
                out int r2, out int c2, out FormulaArgumentValue areaError)) {
                selected = areaError.IsError ? LookupScalarSource(areaError)
                    : new FormulaLookupSource(r2 - r1 + 1, c2 - c1 + 1,
                        (row, column) => NormalizeLookupResult(ResolveFormulaDependency(sheet, r1 + row, c1 + column)));
                return true;
            }
            if (tokens.Count == 4) return false;
            int budget = MaxResolvedFormulaRangeCells;
            if (!TryResolveLookupSource(tokens[0], ref budget, out FormulaLookupSource input)
                || !TrySelectIndexBounds(tokens, input.Rows, input.Columns, out int row, out int column,
                    out int rows, out int columns, out FormulaArgumentValue error)) return false;
            if (input.Error.IsError) error = input.Error;
            selected = error.IsError ? LookupScalarSource(error) : input.Slice(row, column, rows, columns);
            return true;
        }

        private bool TryResolveIndexRange(string args, int? currentRow, out ExcelSheet sheet,
            out int r1, out int c1, out int r2, out int c2, out FormulaArgumentValue error) {
            sheet = this;
            r1 = c1 = r2 = c2 = 0;
            error = default;
            if (_indexFormulaReferenceDepth >= 32 || !HasSufficientFormulaExecutionStack()) return false;
            _indexFormulaReferenceDepth++;
            FormulaArgumentValue previousError = _formulaReferenceError;
            _formulaReferenceError = default;
            try {
                var tokens = SplitFormulaArguments(args, preserveEmpty: true);
                if (tokens.Count < 2 || tokens.Count > 4) return false;
                int area = 1;
                FormulaArgumentValue areaError = default;
                if (tokens.Count == 4 && tokens[3].Length > 0
                    && !TryGetLookupPositionArgument(tokens[3], out area, out areaError)) return false;
                string source = tokens[0].Trim();
                IReadOnlyList<string> areas = source.StartsWith("(", StringComparison.Ordinal) && source.EndsWith(")", StringComparison.Ordinal)
                    ? SplitFormulaArguments(source.Substring(1, source.Length - 2), preserveEmpty: true)
                    : new[] { source };
                int baseR1 = 0, baseC1 = 0, baseR2 = 0, baseC2 = 0;
                ExcelSheet selectedSheet = this;
                ExcelSheet? firstAreaSheet = null;
                bool differentSheets = false;
                for (int index = 0; index < areas.Count; index++) {
                    if (!TryResolveFormulaRangeReference(areas[index], currentRow, out ExcelSheet areaSheet,
                        out int ar1, out int ac1, out int ar2, out int ac2)
                        && !TryParseQualifiedFormulaWholeRange(areas[index], this, out areaSheet,
                            out ar1, out ac1, out ar2, out ac2, out _)
                        && !_formulaReferenceError.IsError) return false;
                    if (firstAreaSheet == null) firstAreaSheet = areaSheet;
                    else if (!ReferenceEquals(firstAreaSheet._worksheetPart, areaSheet._worksheetPart)) differentSheets = true;
                    if (index == area - 1) {
                        selectedSheet = areaSheet;
                        baseR1 = ar1; baseC1 = ac1; baseR2 = ar2; baseC2 = ac2;
                    }
                }
                bool invalidArea = area < 1 || area > areas.Count;
                int sourceRows = invalidArea ? 1 : baseR2 - baseR1 + 1;
                int sourceColumns = invalidArea ? 1 : baseC2 - baseC1 + 1;
                if (!TrySelectIndexBounds(tokens, sourceRows, sourceColumns, out int row, out int column,
                    out int rows, out int columns, out error)) return false;
                if (_formulaReferenceError.IsError) error = _formulaReferenceError;
                else if (differentSheets) error = FormulaArgumentValue.Error("#VALUE!");
                else if (areaError.IsError) error = areaError;
                else if (invalidArea) error = FormulaArgumentValue.Error(area < 1 ? "#VALUE!" : "#REF!");
                if (error.IsError) return true;
                sheet = selectedSheet;
                r1 = baseR1 + row; c1 = baseC1 + column;
                r2 = r1 + rows - 1; c2 = c1 + columns - 1;
                if (!TryGetFormulaRangeCellCount(r1, c1, r2, c2, out _)) return false;
                return true;
            } finally {
                _formulaReferenceError = previousError;
                _indexFormulaReferenceDepth--;
            }
        }

        private bool TrySelectIndexBounds(IReadOnlyList<string> tokens, int sourceRows, int sourceColumns,
            out int row, out int column, out int rows, out int columns, out FormulaArgumentValue error) {
            row = column = rows = columns = 0;
            error = default;
            int rowIndex = 0, columnIndex = 0;
            FormulaArgumentValue rowError = default, columnError = default;
            if (tokens[1].Length > 0 && !TryGetLookupPositionArgument(tokens[1], out rowIndex, out rowError)) return false;
            if (tokens.Count >= 3 && tokens[2].Length > 0
                && !TryGetLookupPositionArgument(tokens[2], out columnIndex, out columnError)) return false;
            if (tokens.Count == 2) {
                if (sourceColumns == 1) columnIndex = 1;
                else if (sourceRows == 1) { columnIndex = rowIndex; rowIndex = 1; }
            }
            if (rowError.IsError) error = rowError;
            else if (columnError.IsError) error = columnError;
            else if (rowIndex < 0 || columnIndex < 0) error = FormulaArgumentValue.Error("#VALUE!");
            else if (rowIndex > sourceRows || columnIndex > sourceColumns) error = FormulaArgumentValue.Error("#REF!");
            if (error.IsError) return true;
            row = rowIndex == 0 ? 0 : rowIndex - 1;
            column = columnIndex == 0 ? 0 : columnIndex - 1;
            rows = rowIndex == 0 ? sourceRows : 1;
            columns = columnIndex == 0 ? sourceColumns : 1;
            return true;
        }
    }
}
