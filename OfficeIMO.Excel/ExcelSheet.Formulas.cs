using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int MaxSupportedFormulaLength = 8192;
        private static readonly TimeSpan FormulaRegexTimeout = TimeSpan.FromMilliseconds(100);

        private static readonly Regex SimpleFunctionFormulaRegex = new Regex(
            @"^\s*=?\s*(SUM|AVERAGE|MIN|MAX|COUNT|COUNTA)\s*\(([^)]*)\)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        private static readonly Regex SimpleBinaryFormulaRegex = new Regex(
            @"^\s*=?\s*([A-Z]+[0-9]+|-?\d+(?:\.\d+)?)\s*([+\-*/])\s*([A-Z]+[0-9]+|-?\d+(?:\.\d+)?)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        /// <summary>
        /// Marks all formula cells on this sheet dirty.
        /// </summary>
        public void InvalidateFormulas() {
            WriteLock(() => {
                foreach (var formula in WorksheetRoot.Descendants<CellFormula>()) {
                    formula.CalculateCell = true;
                }
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Removes cached values from formula cells on this sheet.
        /// </summary>
        public void ClearCachedFormulaResults() {
            WriteLock(() => {
                foreach (var cell in WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula != null)) {
                    cell.CellValue = null;
                }
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Evaluates supported formulas on this sheet and writes cached results.
        /// </summary>
        public int RecalculateSupportedFormulas() {
            int count = 0;
            WriteLock(() => {
                foreach (var cell in WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula != null).ToList()) {
                    if (TryEvaluateFormula(cell.CellFormula!.Text ?? string.Empty, out double result)) {
                        cell.CellValue = new CellValue(result.ToString(CultureInfo.InvariantCulture));
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                        cell.CellFormula.CalculateCell = false;
                        count++;
                    }
                }

                WorksheetRoot.Save();
            });

            return count;
        }

        /// <summary>
        /// Returns the formula text from a cell, if present.
        /// </summary>
        public string? GetFormulaText(int row, int column) {
            return GetCell(row, column).CellFormula?.Text;
        }

        /// <summary>
        /// Tries to return a formula cell's cached value.
        /// </summary>
        public bool TryGetCachedFormulaValue(int row, int column, out string? value) {
            var cell = GetCell(row, column);
            value = cell.CellFormula == null ? null : cell.CellValue?.Text;
            return value != null;
        }

        /// <summary>
        /// Sets a shared-free array formula over a range. The top-left cell owns the formula metadata.
        /// </summary>
        public void SetArrayFormula(string a1Range, string formula) {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentNullException(nameof(formula));
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            WriteLock(() => {
                foreach (var cell in WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula?.FormulaType?.Value == CellFormulaValues.Array).ToList()) {
                    string? reference = cell.CellFormula?.Reference?.Value;
                    if (!string.IsNullOrWhiteSpace(reference)
                        && A1.TryParseRange(reference!, out int existingR1, out int existingC1, out int existingR2, out int existingC2)
                        && RangesOverlapInclusive((r1, c1, r2, c2), (existingR1, existingC1, existingR2, existingC2))) {
                        throw new InvalidOperationException($"Array formula range '{a1Range}' overlaps existing array formula range '{reference}'.");
                    }
                }

                var topLeft = GetCell(r1, c1);
                topLeft.CellFormula = new CellFormula(Utilities.ExcelSanitizer.SanitizeFormula(formula)) {
                    FormulaType = CellFormulaValues.Array,
                    Reference = a1Range
                };
                for (int row = r1; row <= r2; row++) {
                    for (int column = c1; column <= c2; column++) {
                        if (row == r1 && column == c1) continue;
                        var cell = GetCell(row, column);
                        cell.CellFormula = null;
                        cell.CellValue = null;
                    }
                }

                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Clears an array formula whose reference overlaps the supplied range or cell.
        /// </summary>
        public void ClearArrayFormula(string a1RangeOrCell) {
            var bounds = a1RangeOrCell.IndexOf(':') >= 0
                ? A1.ParseRange(a1RangeOrCell)
                : CellAsRange(a1RangeOrCell);

            WriteLock(() => {
                foreach (var cell in WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula?.FormulaType?.Value == CellFormulaValues.Array).ToList()) {
                    string? reference = cell.CellFormula?.Reference?.Value;
                    if (!string.IsNullOrWhiteSpace(reference)
                        && A1.TryParseRange(reference!, out int existingR1, out int existingC1, out int existingR2, out int existingC2)
                        && RangesOverlapInclusive(bounds, (existingR1, existingC1, existingR2, existingC2))) {
                        cell.CellFormula = null;
                        cell.CellValue = null;
                    }
                }

                WorksheetRoot.Save();
            });
        }

        private bool TryEvaluateFormula(string formula, out double result) {
            result = 0;
            if (string.IsNullOrWhiteSpace(formula) || formula.Length > MaxSupportedFormulaLength) {
                return false;
            }

            try {
                var functionMatch = SimpleFunctionFormulaRegex.Match(formula);
                if (functionMatch.Success) {
                    if (!TryResolveFormulaArguments(functionMatch.Groups[2].Value, out var values)) {
                        return false;
                    }

                    string function = functionMatch.Groups[1].Value.ToUpperInvariant();
                    if (function == "COUNTA") {
                        result = values.Count(v => v.HasValue || !string.IsNullOrEmpty(v.Text));
                        return true;
                    }

                    var numbers = values.Where(v => v.Number.HasValue).Select(v => v.Number!.Value).ToList();
                    if (function == "COUNT") {
                        result = numbers.Count;
                        return true;
                    }

                    if (numbers.Count == 0) {
                        return false;
                    }

                    if (function == "SUM") result = numbers.Sum();
                    else if (function == "AVERAGE") result = numbers.Average();
                    else if (function == "MIN") result = numbers.Min();
                    else if (function == "MAX") result = numbers.Max();
                    else return false;
                    return true;
                }

                var binaryMatch = SimpleBinaryFormulaRegex.Match(formula);
                if (binaryMatch.Success) {
                    if (!TryResolveNumericOperand(binaryMatch.Groups[1].Value, out double left)
                        || !TryResolveNumericOperand(binaryMatch.Groups[3].Value, out double right)) {
                        return false;
                    }

                    switch (binaryMatch.Groups[2].Value) {
                        case "+":
                            result = left + right;
                            return true;
                        case "-":
                            result = left - right;
                            return true;
                        case "*":
                            result = left * right;
                            return true;
                        case "/":
                            if (Math.Abs(right) < double.Epsilon) return false;
                            result = left / right;
                            return true;
                    }
                }
            } catch (RegexMatchTimeoutException) {
                return false;
            }

            return false;
        }

        private bool TryResolveFormulaArguments(string args, out List<FormulaArgumentValue> values) {
            values = new List<FormulaArgumentValue>();
            foreach (var token in args.Split(',')) {
                string trimmed = token.Trim();
                if (trimmed.Length == 0) continue;
                if (trimmed.IndexOf(':') >= 0) {
                    if (!A1.TryParseRange(trimmed.Replace("$", string.Empty), out int r1, out int c1, out int r2, out int c2)) {
                        values.Clear();
                        return false;
                    }

                    for (int row = r1; row <= r2; row++) {
                        for (int column = c1; column <= c2; column++) {
                            values.Add(ResolveCellArgument(row, column));
                        }
                    }
                    continue;
                }

                if (TryParseFormulaCellReference(trimmed, out var cellRef)) {
                    values.Add(ResolveCellArgument(cellRef.Row, cellRef.Col));
                    continue;
                }

                if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric)) {
                    values.Add(new FormulaArgumentValue(numeric, trimmed));
                    continue;
                }

                values.Clear();
                return false;
            }

            return true;
        }

        private bool TryResolveNumericOperand(string token, out double value) {
            token = token.Trim().Replace("$", string.Empty);
            if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) {
                return true;
            }

            if (!TryParseFormulaCellReference(token, out var cellRef)) {
                return false;
            }

            var argument = ResolveCellArgument(cellRef.Row, cellRef.Col);
            if (argument.Number.HasValue) {
                value = argument.Number.Value;
                return true;
            }

            return false;
        }

        private static bool TryParseFormulaCellReference(string token, out (int Row, int Col) cellRef) {
            cellRef = A1.ParseCellRef(token.Trim().Replace("$", string.Empty));
            return cellRef.Row > 0
                && cellRef.Col > 0
                && cellRef.Row <= A1.MaxRows
                && cellRef.Col <= A1.MaxColumns;
        }

        private FormulaArgumentValue ResolveCellArgument(int row, int column) {
            var value = GetCellValueSnapshot(row, column);
            if (value.Value is double d) {
                return new FormulaArgumentValue(d, value.CachedText);
            }

            if (double.TryParse(value.CachedText, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) {
                return new FormulaArgumentValue(parsed, value.CachedText);
            }

            return new FormulaArgumentValue(null, value.Value?.ToString());
        }

        private readonly struct FormulaArgumentValue {
            internal FormulaArgumentValue(double? number, string? text) {
                Number = number;
                Text = text;
            }

            internal double? Number { get; }
            internal string? Text { get; }
            internal bool HasValue => Number.HasValue || Text != null;
        }
    }
}
