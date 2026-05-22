using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int MaxSupportedFormulaLength = 8192;
        private static readonly TimeSpan FormulaRegexTimeout = TimeSpan.FromMilliseconds(100);
        private Dictionary<string, FormulaArgumentValue>? _formulaEvaluationCache;
        private HashSet<string>? _formulaEvaluationStack;

        private static readonly Regex SimpleFunctionFormulaRegex = new Regex(
            @"^\s*=?\s*(SUM|AVERAGE|MIN|MAX|COUNT|COUNTA|COUNTIF|SUMIF|AVERAGEIF|COUNTIFS|SUMIFS|AVERAGEIFS|PRODUCT|MEDIAN|LARGE|SMALL|SUMPRODUCT|VLOOKUP|HLOOKUP|XLOOKUP|ABS|SIGN|ROUND|ROUNDUP|ROUNDDOWN|TRUNC|INT|CEILING|FLOOR|POWER|SQRT|LN|LOG10|EXP|PI|RADIANS|DEGREES|MOD|DATE|TIME|TODAY|NOW|YEAR|MONTH|DAY|HOUR|MINUTE|SECOND|EDATE|EOMONTH|DAYS|WEEKDAY|NETWORKDAYS|IF|AND|OR|NOT|IFERROR|CONCAT|TEXTJOIN|LEFT|RIGHT|MID|LEN|TRIM)\s*\((.*)\)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        private static readonly Regex FunctionNameFormulaRegex = new Regex(
            @"^\s*=?\s*([A-Za-z][A-Za-z0-9_.]*)\s*\(",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        private static readonly Regex SimpleBinaryFormulaRegex = new Regex(
            @"^\s*=?\s*((?:'(?:[^']|'')+'|[A-Za-z_][^!+\-*/<>=,\(\)]*)!(?:\$?[A-Z]+\$?[0-9]+|[A-Za-z_][A-Za-z0-9_.]*(?:\[[^+\-*/<>=\(\)]*\])*)|\$?[A-Z]+\$?[0-9]+|[A-Za-z_][A-Za-z0-9_.]*(?:\[[^+\-*/<>=\(\)]*\])*|-?\d+(?:\.\d+)?)\s*([+\-*/])\s*((?:'(?:[^']|'')+'|[A-Za-z_][^!+\-*/<>=,\(\)]*)!(?:\$?[A-Z]+\$?[0-9]+|[A-Za-z_][A-Za-z0-9_.]*(?:\[[^+\-*/<>=\(\)]*\])*)|\$?[A-Z]+\$?[0-9]+|[A-Za-z_][A-Za-z0-9_.]*(?:\[[^+\-*/<>=\(\)]*\])*|-?\d+(?:\.\d+)?)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        private static readonly Regex SimpleComparisonFormulaRegex = new Regex(
            @"^\s*((?:'(?:[^']|'')+'|[A-Za-z_][^!+\-*/<>=,\(\)]*)!(?:\$?[A-Z]+\$?[0-9]+|[A-Za-z_][A-Za-z0-9_.]*(?:\[[^+\-*/<>=\(\)]*\])*)|\$?[A-Z]+\$?[0-9]+|[A-Za-z_][A-Za-z0-9_.]*(?:\[[^+\-*/<>=\(\)]*\])*|-?\d+(?:\.\d+)?)\s*(>=|<=|<>|=|>|<)\s*((?:'(?:[^']|'')+'|[A-Za-z_][^!+\-*/<>=,\(\)]*)!(?:\$?[A-Z]+\$?[0-9]+|[A-Za-z_][A-Za-z0-9_.]*(?:\[[^+\-*/<>=\(\)]*\])*)|\$?[A-Z]+\$?[0-9]+|[A-Za-z_][A-Za-z0-9_.]*(?:\[[^+\-*/<>=\(\)]*\])*|-?\d+(?:\.\d+)?)\s*$",
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
                bool changed = false;
                foreach (var cell in WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula != null)) {
                    if (cell.CellValue != null) {
                        cell.CellValue = null;
                        changed = true;
                    }
                }

                if (changed) {
                    _hasWorksheetMutations = true;
                    MarkRequiresSavePreparation();
                    ClearCellTextSharedStringCache();
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
                var previousCache = _formulaEvaluationCache;
                var previousStack = _formulaEvaluationStack;
                _formulaEvaluationCache = new Dictionary<string, FormulaArgumentValue>(StringComparer.OrdinalIgnoreCase);
                _formulaEvaluationStack = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                try {
                    foreach (var cell in WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula != null).ToList()) {
                        if (!TryEvaluateFormulaCellValue(cell, out FormulaArgumentValue result)) {
                            continue;
                        }

                        SetFormulaCachedValue(cell, result);
                        cell.CellFormula!.CalculateCell = false;
                        count++;
                    }
                } finally {
                    _formulaEvaluationCache = previousCache;
                    _formulaEvaluationStack = previousStack;
                }

                if (count > 0) {
                    _hasWorksheetMutations = true;
                    MarkRequiresSavePreparation();
                    ClearCellTextSharedStringCache();
                }

                WorksheetRoot.Save();
            });

            return count;
        }

        private bool TryEvaluateFormulaCell(Cell cell, out double result) {
            result = 0;
            if (!TryEvaluateFormulaCellValue(cell, out FormulaArgumentValue value) || !value.Number.HasValue) {
                return false;
            }

            result = value.Number.Value;
            return true;
        }

        private bool TryEvaluateFormulaCellValue(Cell cell, out FormulaArgumentValue result) {
            result = default;
            if (cell.CellFormula == null) {
                return false;
            }

            string? reference = NormalizeFormulaCellReference(cell.CellReference?.Value);
            if (reference == null || _formulaEvaluationCache == null || _formulaEvaluationStack == null) {
                return TryEvaluateFormulaValue(cell.CellFormula.Text ?? string.Empty, out result);
            }

            string cacheKey = GetFormulaEvaluationCacheKey(reference);
            if (_formulaEvaluationCache.TryGetValue(cacheKey, out result)) {
                return true;
            }

            if (!_formulaEvaluationStack.Add(cacheKey)) {
                return false;
            }

            try {
                if (!TryEvaluateFormulaValue(cell.CellFormula.Text ?? string.Empty, out result)) {
                    return false;
                }

                _formulaEvaluationCache[cacheKey] = result;
                return true;
            } finally {
                _formulaEvaluationStack.Remove(cacheKey);
            }
        }

        private static void SetFormulaCachedValue(Cell cell, FormulaArgumentValue result) {
            if (result.Number.HasValue) {
                cell.CellValue = new CellValue(result.Number.Value.ToString(CultureInfo.InvariantCulture));
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                return;
            }

            if (result.Text != null) {
                cell.CellValue = new CellValue(result.Text);
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            }
        }

        /// <summary>
        /// Inspects formula cells on this sheet without changing workbook contents.
        /// </summary>
        public ExcelFormulaInspection InspectFormulas() {
            return new ExcelFormulaInspection(GetFormulaCells());
        }

        /// <summary>
        /// Returns formula cells on this sheet without changing workbook contents.
        /// </summary>
        public IReadOnlyList<ExcelFormulaCellInfo> GetFormulaCells() {
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                var formulas = new List<ExcelFormulaCellInfo>();
                foreach (var cell in WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula != null)) {
                    string formula = cell.CellFormula!.Text ?? string.Empty;
                    bool supported = TryEvaluateFormulaValue(formula, out _);
                    formulas.Add(new ExcelFormulaCellInfo(
                        Name,
                        cell.CellReference?.Value ?? string.Empty,
                        formula,
                        cell.CellValue?.Text,
                        cell.CellFormula.CalculateCell?.Value ?? false,
                        supported,
                        supported ? null : GetUnsupportedFormulaReason(formula)));
                }

                return formulas;
            });
        }

        /// <summary>
        /// Returns the formula text from a cell, if present.
        /// </summary>
        public string? GetFormulaText(int row, int column) {
            return TryGetExistingCell(row, column)?.CellFormula?.Text;
        }

        /// <summary>
        /// Tries to return a formula cell's cached value.
        /// </summary>
        public bool TryGetCachedFormulaValue(int row, int column, out string? value) {
            var cell = TryGetExistingCell(row, column);
            value = cell?.CellFormula == null ? null : cell.CellValue?.Text;
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
                        && A1.TryParseRange(reference!.Replace("$", string.Empty), out int existingR1, out int existingC1, out int existingR2, out int existingC2)
                        && RangesOverlapInclusive(bounds, (existingR1, existingC1, existingR2, existingC2))) {
                        for (int row = existingR1; row <= existingR2; row++) {
                            for (int column = existingC1; column <= existingC2; column++) {
                                var spillCell = TryGetExistingCell(row, column);
                                if (spillCell == null) {
                                    continue;
                                }

                                spillCell.CellFormula = null;
                                spillCell.CellValue = null;
                            }
                        }
                    }
                }

                WorksheetRoot.Save();
            });
        }

        private bool TryEvaluateFormulaValue(string formula, out FormulaArgumentValue result) {
            result = default;
            if (string.IsNullOrWhiteSpace(formula) || formula.Length > MaxSupportedFormulaLength) {
                return false;
            }

            try {
                var functionMatch = SimpleFunctionFormulaRegex.Match(formula);
                if (functionMatch.Success) {
                    string function = functionMatch.Groups[1].Value.ToUpperInvariant();
                    string args = functionMatch.Groups[2].Value;
                    if (TryEvaluateTextFunction(function, args, out result)) {
                        return true;
                    }

                    if ((function == "VLOOKUP" || function == "HLOOKUP" || function == "XLOOKUP")
                        && TryEvaluateLookupValue(function, args, out result)) {
                        return true;
                    }
                }
            } catch (RegexMatchTimeoutException) {
                return false;
            }

            if (TryEvaluateFormula(formula, out double numeric)) {
                result = new FormulaArgumentValue(numeric, numeric.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            return false;
        }

        private bool TryEvaluateFormula(string formula, out double result) {
            result = 0;
            if (string.IsNullOrWhiteSpace(formula) || formula.Length > MaxSupportedFormulaLength) {
                return false;
            }

            try {
                var functionMatch = SimpleFunctionFormulaRegex.Match(formula);
                if (functionMatch.Success) {
                    string function = functionMatch.Groups[1].Value.ToUpperInvariant();
                    string args = functionMatch.Groups[2].Value;
                    if (function == "IFERROR") {
                        if (!TryEvaluateIfError(args, out result)) {
                            return false;
                        }

                        return true;
                    }

                    if (function == "IF") {
                        if (!TryEvaluateIf(args, out result)) {
                            return false;
                        }

                        return true;
                    }

                    if (function == "AND" || function == "OR") {
                        if (!TryEvaluateLogical(args, useAnd: function == "AND", out bool logicalResult)) {
                            return false;
                        }

                        result = logicalResult ? 1d : 0d;
                        return true;
                    }

                    if (function == "NOT") {
                        if (!TryEvaluateNot(args, out bool logicalResult)) {
                            return false;
                        }

                        result = logicalResult ? 1d : 0d;
                        return true;
                    }

                    if (function == "COUNTIF" || function == "SUMIF" || function == "AVERAGEIF") {
                        return TryEvaluateConditionalAggregate(function, args, out result);
                    }

                    if (function == "COUNTIFS" || function == "SUMIFS" || function == "AVERAGEIFS") {
                        return TryEvaluateMultiCriteriaAggregate(function, args, out result);
                    }

                    if (function == "DATE" || function == "TIME" || function == "TODAY" || function == "NOW"
                        || function == "YEAR" || function == "MONTH" || function == "DAY"
                        || function == "HOUR" || function == "MINUTE" || function == "SECOND"
                        || function == "EDATE" || function == "EOMONTH" || function == "DAYS"
                        || function == "WEEKDAY" || function == "NETWORKDAYS") {
                        return TryEvaluateDateTimeFunction(function, args, out result);
                    }

                    if (function == "SUMPRODUCT") {
                        return TryEvaluateSumProduct(args, out result);
                    }

                    if (function == "VLOOKUP" || function == "HLOOKUP" || function == "XLOOKUP") {
                        return TryEvaluateLookupFunction(function, args, out result);
                    }

                    if (!TryResolveFormulaArguments(args, out var values) || values.Any(value => value.IsUnresolvedFormula)) {
                        return false;
                    }

                    if (function == "COUNTA") {
                        result = values.Count(v => v.HasValue || !string.IsNullOrEmpty(v.Text));
                        return true;
                    }

                    var numbers = values.Where(v => v.Number.HasValue).Select(v => v.Number!.Value).ToList();
                    if (function == "COUNT") {
                        result = numbers.Count;
                        return true;
                    }

                    if (function == "ABS") {
                        if (numbers.Count != 1) {
                            return false;
                        }

                        result = Math.Abs(numbers[0]);
                        return true;
                    }

                    if (function == "SIGN") {
                        if (numbers.Count != 1) {
                            return false;
                        }

                        result = Math.Sign(numbers[0]);
                        return true;
                    }

                    if (function == "ROUND") {
                        if (numbers.Count != 2) {
                            return false;
                        }

                        int digits = (int)Math.Round(numbers[1], MidpointRounding.AwayFromZero);
                        result = Math.Round(numbers[0], digits, MidpointRounding.AwayFromZero);
                        return true;
                    }

                    if (function == "ROUNDUP" || function == "ROUNDDOWN") {
                        if (numbers.Count != 2 || !TryGetSupportedDecimalPlaces(numbers[1], out int digits)) {
                            return false;
                        }

                        double factor = Math.Pow(10, digits);
                        double shifted = Math.Abs(numbers[0]) * factor;
                        double rounded = function == "ROUNDUP" ? Math.Ceiling(shifted) : Math.Floor(shifted);
                        result = Math.Sign(numbers[0]) * rounded / factor;
                        return true;
                    }

                    if (function == "TRUNC") {
                        if (numbers.Count < 1 || numbers.Count > 2) {
                            return false;
                        }

                        int digits = 0;
                        if (numbers.Count == 2 && !TryGetSupportedDecimalPlaces(numbers[1], out digits)) {
                            return false;
                        }

                        double factor = Math.Pow(10, digits);
                        result = Math.Truncate(numbers[0] * factor) / factor;
                        return true;
                    }

                    if (function == "INT") {
                        if (numbers.Count != 1) {
                            return false;
                        }

                        result = Math.Floor(numbers[0]);
                        return true;
                    }

                    if (function == "CEILING" || function == "FLOOR") {
                        if (numbers.Count != 2 || numbers[1] <= 0) {
                            return false;
                        }

                        double value = numbers[0] / numbers[1];
                        double rounded = function == "CEILING" ? Math.Ceiling(value) : Math.Floor(value);
                        result = rounded * numbers[1];
                        return true;
                    }

                    if (function == "POWER") {
                        if (numbers.Count != 2) {
                            return false;
                        }

                        double value = Math.Pow(numbers[0], numbers[1]);
                        if (double.IsNaN(value) || double.IsInfinity(value)) {
                            return false;
                        }

                        result = value;
                        return true;
                    }

                    if (function == "SQRT") {
                        if (numbers.Count != 1 || numbers[0] < 0) {
                            return false;
                        }

                        result = Math.Sqrt(numbers[0]);
                        return true;
                    }

                    if (function == "LN" || function == "LOG10") {
                        if (numbers.Count != 1 || numbers[0] <= 0) {
                            return false;
                        }

                        result = function == "LN" ? Math.Log(numbers[0]) : Math.Log10(numbers[0]);
                        return true;
                    }

                    if (function == "EXP") {
                        if (numbers.Count != 1) {
                            return false;
                        }

                        result = Math.Exp(numbers[0]);
                        if (double.IsNaN(result) || double.IsInfinity(result)) {
                            return false;
                        }

                        return true;
                    }

                    if (function == "PI") {
                        if (numbers.Count != 0) {
                            return false;
                        }

                        result = Math.PI;
                        return true;
                    }

                    if (function == "RADIANS" || function == "DEGREES") {
                        if (numbers.Count != 1) {
                            return false;
                        }

                        result = function == "RADIANS" ? numbers[0] * Math.PI / 180d : numbers[0] * 180d / Math.PI;
                        return true;
                    }

                    if (function == "MOD") {
                        if (numbers.Count != 2 || Math.Abs(numbers[1]) < double.Epsilon) {
                            return false;
                        }

                        result = numbers[0] - numbers[1] * Math.Floor(numbers[0] / numbers[1]);
                        return true;
                    }

                    if (function == "LARGE" || function == "SMALL") {
                        if (numbers.Count < 2 || !TryGetWholeNumber(numbers[numbers.Count - 1], out int rank)) {
                            return false;
                        }

                        var sorted = numbers.Take(numbers.Count - 1).OrderBy(value => value).ToList();
                        if (rank < 1 || rank > sorted.Count) {
                            return false;
                        }

                        result = function == "LARGE" ? sorted[sorted.Count - rank] : sorted[rank - 1];
                        return true;
                    }

                    if (numbers.Count == 0) {
                        return false;
                    }

                    if (function == "SUM") result = numbers.Sum();
                    else if (function == "AVERAGE") result = numbers.Average();
                    else if (function == "MIN") result = numbers.Min();
                    else if (function == "MAX") result = numbers.Max();
                    else if (function == "PRODUCT") result = numbers.Aggregate(1d, (current, value) => current * value);
                    else if (function == "MEDIAN") result = CalculateMedian(numbers);
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

        private bool TryEvaluateTextFunction(string function, string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (function == "CONCAT") {
                if (tokens.Count == 0 || !TryResolveTextArgumentValues(tokens, out var parts)) {
                    return false;
                }

                result = new FormulaArgumentValue(null, string.Concat(parts));
                return true;
            }

            if (function == "TEXTJOIN") {
                if (tokens.Count < 3
                    || !TryResolveTextArgument(tokens[0], out string delimiter)
                    || !TryResolveBooleanArgument(tokens[1], out bool ignoreEmpty)
                    || !TryResolveTextArgumentValues(tokens.Skip(2), out var parts)) {
                    return false;
                }

                if (ignoreEmpty) {
                    parts = parts.Where(part => part.Length > 0).ToList();
                }

                result = new FormulaArgumentValue(null, string.Join(delimiter, parts));
                return true;
            }

            if (function == "LEFT" || function == "RIGHT") {
                if (tokens.Count < 1
                    || tokens.Count > 2
                    || !TryResolveTextArgument(tokens[0], out string text)) {
                    return false;
                }

                int count = 1;
                if (tokens.Count == 2 && !TryGetWholeNumberArgument(tokens[1], out count)) {
                    return false;
                }

                if (count < 0) {
                    return false;
                }

                count = Math.Min(count, text.Length);
                result = new FormulaArgumentValue(null, function == "LEFT"
                    ? text.Substring(0, count)
                    : text.Substring(text.Length - count, count));
                return true;
            }

            if (function == "MID") {
                if (tokens.Count != 3
                    || !TryResolveTextArgument(tokens[0], out string text)
                    || !TryGetWholeNumberArgument(tokens[1], out int start)
                    || !TryGetWholeNumberArgument(tokens[2], out int count)
                    || start < 1
                    || count < 0) {
                    return false;
                }

                int startIndex = start - 1;
                if (startIndex >= text.Length) {
                    result = new FormulaArgumentValue(null, string.Empty);
                    return true;
                }

                count = Math.Min(count, text.Length - startIndex);
                result = new FormulaArgumentValue(null, text.Substring(startIndex, count));
                return true;
            }

            if (function == "LEN") {
                if (tokens.Count != 1 || !TryResolveTextArgument(tokens[0], out string text)) {
                    return false;
                }

                result = new FormulaArgumentValue(text.Length, text.Length.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (function == "TRIM") {
                if (tokens.Count != 1 || !TryResolveTextArgument(tokens[0], out string text)) {
                    return false;
                }

                result = new FormulaArgumentValue(null, string.Join(" ", text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)));
                return true;
            }

            return false;
        }

        private bool TryEvaluateLookupFunction(string function, string args, out double result) {
            result = 0;
            if (!TryEvaluateLookupValue(function, args, out FormulaArgumentValue value) || !value.Number.HasValue) {
                return false;
            }

            result = value.Number.Value;
            return true;
        }

        private bool TryEvaluateLookupValue(string function, string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (function == "XLOOKUP") {
                return TryEvaluateXLookupValue(tokens, out result);
            }

            if (tokens.Count != 4
                || !TryResolveFormulaArgument(tokens[0], out var lookupValue)
                || !TryGetWholeNumberArgument(tokens[2], out int resultIndex)
                || !IsExactLookupMode(tokens[3])
                || !TryResolveFormulaRangeReference(tokens[1], out ExcelSheet rangeSheet, out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }

            if (function == "VLOOKUP") {
                int width = c2 - c1 + 1;
                if (resultIndex < 1 || resultIndex > width) {
                    return false;
                }

                int resultColumn = c1 + resultIndex - 1;
                for (int row = r1; row <= r2; row++) {
                    if (!FormulaValuesEqual(rangeSheet.ResolveCellArgument(row, c1), lookupValue)) {
                        continue;
                    }

                    result = rangeSheet.ResolveCellArgument(row, resultColumn);
                    return result.HasValue;
                }

                return false;
            }

            int height = r2 - r1 + 1;
            if (resultIndex < 1 || resultIndex > height) {
                return false;
            }

            int resultRow = r1 + resultIndex - 1;
            for (int column = c1; column <= c2; column++) {
                if (!FormulaValuesEqual(rangeSheet.ResolveCellArgument(r1, column), lookupValue)) {
                    continue;
                }

                result = rangeSheet.ResolveCellArgument(resultRow, column);
                return result.HasValue;
            }

            return false;
        }

        private bool TryEvaluateXLookupValue(IReadOnlyList<string> tokens, out FormulaArgumentValue result) {
            result = default;
            if (tokens.Count != 3
                || !TryResolveFormulaArgument(tokens[0], out var lookupValue)
                || !TryResolveFormulaRange(tokens[1], out var lookupValues)
                || !TryResolveFormulaRange(tokens[2], out var returnValues)
                || lookupValues.Count != returnValues.Count) {
                return false;
            }

            for (int index = 0; index < lookupValues.Count; index++) {
                if (!FormulaValuesEqual(lookupValues[index], lookupValue)) {
                    continue;
                }

                var returnValue = returnValues[index];
                result = returnValue;
                return result.HasValue;
            }

            return false;
        }

        private bool TryEvaluateSumProduct(string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count == 0) {
                return false;
            }

            var argumentSets = new List<List<double>>();
            foreach (string token in tokens) {
                if (!TryResolveFormulaArgumentNumbers(token, out var values) || values.Count == 0) {
                    return false;
                }

                if (argumentSets.Count > 0 && values.Count != argumentSets[0].Count) {
                    return false;
                }

                argumentSets.Add(values);
            }

            for (int index = 0; index < argumentSets[0].Count; index++) {
                double product = 1d;
                foreach (var values in argumentSets) {
                    product *= values[index];
                }

                result += product;
            }

            return true;
        }

        private bool TryEvaluateIfError(string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 2) {
                return false;
            }

            if (TryEvaluateFormulaOrNumeric(tokens[0], out result)) {
                return true;
            }

            return TryEvaluateFormulaOrNumeric(tokens[1], out result);
        }

        private bool TryEvaluateIf(string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 3
                || !TryEvaluateCondition(tokens[0], out bool condition)
                || !TryEvaluateFormulaOrNumeric(tokens[condition ? 1 : 2], out result)) {
                return false;
            }

            return true;
        }

        private bool TryEvaluateFormulaOrNumeric(string token, out double result) {
            return TryEvaluateFormula(token, out result) || TryResolveNumericOperand(token, out result);
        }

        private bool TryEvaluateLogical(string args, bool useAnd, out bool result) {
            result = useAnd;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count == 0) {
                return false;
            }

            foreach (string token in tokens) {
                if (!TryEvaluateCondition(token, out bool condition)) {
                    return false;
                }

                if (useAnd && !condition) {
                    result = false;
                    return true;
                }

                if (!useAnd && condition) {
                    result = true;
                    return true;
                }
            }

            result = useAnd;
            return true;
        }

        private bool TryEvaluateNot(string args, out bool result) {
            result = false;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 1 || !TryEvaluateCondition(tokens[0], out bool condition)) {
                return false;
            }

            result = !condition;
            return true;
        }

        private bool TryEvaluateConditionalAggregate(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 2 || tokens.Count > 3
                || !TryResolveFormulaRange(tokens[0], out var criteriaValues)
                || !TryParseCriteria(tokens[1], out var criteria)) {
                return false;
            }

            var aggregateValues = criteriaValues;
            if (tokens.Count == 3) {
                if (!TryResolveFormulaRange(tokens[2], out aggregateValues) || aggregateValues.Count != criteriaValues.Count) {
                    return false;
                }
            } else if (function == "AVERAGEIF") {
                aggregateValues = criteriaValues;
            }

            var matched = new List<FormulaArgumentValue>();
            for (int index = 0; index < criteriaValues.Count; index++) {
                if (MatchesCriteria(criteriaValues[index], criteria)) {
                    matched.Add(aggregateValues[index]);
                }
            }

            if (function == "COUNTIF") {
                result = matched.Count;
                return true;
            }

            var numbers = matched.Where(value => value.Number.HasValue).Select(value => value.Number!.Value).ToList();
            if (function == "SUMIF") {
                result = numbers.Sum();
                return true;
            }

            if (numbers.Count == 0) {
                return false;
            }

            result = numbers.Average();
            return true;
        }

        private bool TryEvaluateDateTimeFunction(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);

            if (function == "TODAY") {
                if (tokens.Count != 0) {
                    return false;
                }

                result = DateTime.Today.ToOADate();
                return true;
            }

            if (function == "NOW") {
                if (tokens.Count != 0) {
                    return false;
                }

                result = DateTime.Now.ToOADate();
                return true;
            }

            if (function == "NETWORKDAYS") {
                return TryEvaluateNetworkDays(tokens, out result);
            }

            if (!TryResolveFormulaOrNumericArguments(tokens, out var numbers)) {
                return false;
            }

            if (function == "DATE") {
                if (numbers.Count != 3
                    || !TryGetWholeNumber(numbers[0], out int year)
                    || !TryGetWholeNumber(numbers[1], out int month)
                    || !TryGetWholeNumber(numbers[2], out int day)) {
                    return false;
                }

                if (year >= 0 && year <= 1899) {
                    year += 1900;
                }

                if (year < 1 || year > 9999) {
                    return false;
                }

                try {
                    result = new DateTime(year, 1, 1).AddMonths(month - 1).AddDays(day - 1).ToOADate();
                } catch (ArgumentOutOfRangeException) {
                    return false;
                }

                return true;
            }

            if (function == "TIME") {
                if (numbers.Count != 3) {
                    return false;
                }

                double seconds = numbers[0] * 3600d + numbers[1] * 60d + numbers[2];
                seconds %= 86400d;
                if (seconds < 0) {
                    seconds += 86400d;
                }

                result = seconds / 86400d;
                return true;
            }

            if (function == "EDATE" || function == "EOMONTH") {
                if (numbers.Count != 2 || !TryGetWholeNumber(numbers[1], out int months)) {
                    return false;
                }

                if (!TryGetDateFromSerial(numbers[0], out DateTime startDate)) {
                    return false;
                }

                try {
                    DateTime shifted = startDate.AddMonths(months);
                    result = function == "EOMONTH"
                        ? new DateTime(shifted.Year, shifted.Month, DateTime.DaysInMonth(shifted.Year, shifted.Month)).ToOADate()
                        : shifted.ToOADate();
                } catch (ArgumentOutOfRangeException) {
                    return false;
                }

                return true;
            }

            if (function == "DAYS") {
                if (numbers.Count != 2
                    || !TryGetDateFromSerial(numbers[0], out DateTime endDate)
                    || !TryGetDateFromSerial(numbers[1], out DateTime startDate)) {
                    return false;
                }

                result = (endDate - startDate).TotalDays;
                return true;
            }

            if (function == "WEEKDAY") {
                if (numbers.Count < 1 || numbers.Count > 2
                    || !TryGetDateFromSerial(numbers[0], out DateTime date)) {
                    return false;
                }

                int returnType = 1;
                if (numbers.Count == 2 && !TryGetWholeNumber(numbers[1], out returnType)) {
                    return false;
                }

                int day = (int)date.DayOfWeek;
                if (returnType == 1) {
                    result = day + 1;
                    return true;
                }

                if (returnType == 2) {
                    result = day == 0 ? 7 : day;
                    return true;
                }

                if (returnType == 3) {
                    result = day == 0 ? 6 : day - 1;
                    return true;
                }

                return false;
            }

            if (numbers.Count != 1) {
                return false;
            }

            DateTime dateTime;
            try {
                dateTime = DateTime.FromOADate(numbers[0]);
            } catch (ArgumentException) {
                return false;
            }

            switch (function) {
                case "YEAR":
                    result = dateTime.Year;
                    return true;
                case "MONTH":
                    result = dateTime.Month;
                    return true;
                case "DAY":
                    result = dateTime.Day;
                    return true;
                case "HOUR":
                    result = dateTime.Hour;
                    return true;
                case "MINUTE":
                    result = dateTime.Minute;
                    return true;
                case "SECOND":
                    result = dateTime.Second;
                    return true;
                default:
                    return false;
            }
        }

        private bool TryResolveFormulaOrNumericArguments(IReadOnlyList<string> tokens, out List<double> numbers) {
            numbers = new List<double>();
            foreach (string token in tokens) {
                if (!TryEvaluateFormulaOrNumeric(token, out double value)) {
                    return false;
                }

                numbers.Add(value);
            }

            return true;
        }

        private bool TryEvaluateNetworkDays(IReadOnlyList<string> tokens, out double result) {
            result = 0;
            if (tokens.Count < 2 || tokens.Count > 3
                || !TryEvaluateFormulaOrNumeric(tokens[0], out double startSerial)
                || !TryEvaluateFormulaOrNumeric(tokens[1], out double endSerial)
                || !TryGetDateFromSerial(startSerial, out DateTime startDate)
                || !TryGetDateFromSerial(endSerial, out DateTime endDate)) {
                return false;
            }

            var holidays = new HashSet<DateTime>();
            if (tokens.Count == 3 && !TryResolveHolidayDates(tokens[2], holidays)) {
                return false;
            }

            int direction = startDate <= endDate ? 1 : -1;
            DateTime current = direction == 1 ? startDate : endDate;
            DateTime last = direction == 1 ? endDate : startDate;
            int days = 0;
            while (current <= last) {
                if (current.DayOfWeek != DayOfWeek.Saturday
                    && current.DayOfWeek != DayOfWeek.Sunday
                    && !holidays.Contains(current.Date)) {
                    days++;
                }

                current = current.AddDays(1);
            }

            result = days * direction;
            return true;
        }

        private bool TryResolveHolidayDates(string token, HashSet<DateTime> holidays) {
            List<FormulaArgumentValue> values;
            if (token.IndexOf(':') >= 0) {
                if (!TryResolveFormulaRange(token, out values)) {
                    return false;
                }
            } else if (!TryResolveFormulaArguments(token, out values)) {
                return false;
            }

            foreach (var value in values) {
                if (value.Number.HasValue && TryGetDateFromSerial(value.Number.Value, out DateTime date)) {
                    holidays.Add(date);
                }
            }

            return true;
        }

        private bool TryResolveFormulaArgumentNumbers(string token, out List<double> numbers) {
            numbers = new List<double>();
            if (TryResolveFormulaRange(token, out var values)) {
                foreach (var value in values) {
                    if (!value.Number.HasValue) {
                        return false;
                    }

                    numbers.Add(value.Number.Value);
                }

                return true;
            }

            if (!TryEvaluateFormulaOrNumeric(token, out double numeric)) {
                return false;
            }

            numbers.Add(numeric);
            return true;
        }

        private bool TryEvaluateMultiCriteriaAggregate(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            bool countOnly = function == "COUNTIFS";
            if (countOnly) {
                if (tokens.Count < 2 || tokens.Count % 2 != 0) {
                    return false;
                }
            } else if (tokens.Count < 3 || tokens.Count % 2 != 1) {
                return false;
            }

            int criteriaStart = countOnly ? 0 : 1;
            List<FormulaArgumentValue>? aggregateValues = null;
            if (!countOnly && !TryResolveFormulaRange(tokens[0], out aggregateValues)) {
                return false;
            }

            var criteriaSets = new List<(List<FormulaArgumentValue> Values, FormulaCriteria Criteria)>();
            int expectedCount = aggregateValues?.Count ?? -1;
            for (int index = criteriaStart; index < tokens.Count; index += 2) {
                if (!TryResolveFormulaRange(tokens[index], out var criteriaValues)
                    || !TryParseCriteria(tokens[index + 1], out var criteria)) {
                    return false;
                }

                if (expectedCount < 0) {
                    expectedCount = criteriaValues.Count;
                } else if (criteriaValues.Count != expectedCount) {
                    return false;
                }

                criteriaSets.Add((criteriaValues, criteria));
            }

            var matched = new List<FormulaArgumentValue>();
            for (int rowIndex = 0; rowIndex < expectedCount; rowIndex++) {
                bool matches = true;
                foreach (var criteriaSet in criteriaSets) {
                    if (!MatchesCriteria(criteriaSet.Values[rowIndex], criteriaSet.Criteria)) {
                        matches = false;
                        break;
                    }
                }

                if (matches) {
                    matched.Add(countOnly ? criteriaSets[0].Values[rowIndex] : aggregateValues![rowIndex]);
                }
            }

            if (countOnly) {
                result = matched.Count;
                return true;
            }

            var numbers = matched.Where(value => value.Number.HasValue).Select(value => value.Number!.Value).ToList();
            if (function == "SUMIFS") {
                result = numbers.Sum();
                return true;
            }

            if (numbers.Count == 0) {
                return false;
            }

            result = numbers.Average();
            return true;
        }

        private bool TryEvaluateCondition(string condition, out bool result) {
            result = false;
            var comparison = SimpleComparisonFormulaRegex.Match(condition);
            if (comparison.Success) {
                if (!TryResolveNumericOperand(comparison.Groups[1].Value, out double left)
                    || !TryResolveNumericOperand(comparison.Groups[3].Value, out double right)) {
                    return false;
                }

                switch (comparison.Groups[2].Value) {
                    case ">":
                        result = left > right;
                        return true;
                    case "<":
                        result = left < right;
                        return true;
                    case ">=":
                        result = left >= right;
                        return true;
                    case "<=":
                        result = left <= right;
                        return true;
                    case "=":
                        result = Math.Abs(left - right) < 0.0000001;
                        return true;
                    case "<>":
                        result = Math.Abs(left - right) >= 0.0000001;
                        return true;
                }
            }

            if (TryEvaluateFormula(condition, out double formulaValue)) {
                result = Math.Abs(formulaValue) >= double.Epsilon;
                return true;
            }

            if (!TryResolveNumericOperand(condition, out double value)) {
                return false;
            }

            result = Math.Abs(value) >= double.Epsilon;
            return true;
        }

        private bool TryResolveFormulaRange(string token, out List<FormulaArgumentValue> values) {
            values = new List<FormulaArgumentValue>();
            if (!TryResolveFormulaRangeReference(token, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }

            for (int row = r1; row <= r2; row++) {
                for (int column = c1; column <= c2; column++) {
                    values.Add(sheet.ResolveCellArgument(row, column));
                }
            }

            return true;
        }

        private static bool TryParseCriteria(string token, out FormulaCriteria criteria) {
            string value = token.Trim();
            if (value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"') {
                value = value.Substring(1, value.Length - 2);
            }

            string op = "=";
            foreach (string candidate in new[] { ">=", "<=", "<>", ">", "<", "=" }) {
                if (value.StartsWith(candidate, StringComparison.Ordinal)) {
                    op = candidate;
                    value = value.Substring(candidate.Length);
                    break;
                }
            }

            value = value.Trim();
            double? number = double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric)
                ? numeric
                : (double?)null;

            criteria = new FormulaCriteria(op, value, number);
            return value.Length > 0 || op == "=";
        }

        private static bool MatchesCriteria(FormulaArgumentValue value, FormulaCriteria criteria) {
            if (criteria.Number.HasValue && value.Number.HasValue) {
                double left = value.Number.Value;
                double right = criteria.Number.Value;
                switch (criteria.Operator) {
                    case ">":
                        return left > right;
                    case "<":
                        return left < right;
                    case ">=":
                        return left >= right;
                    case "<=":
                        return left <= right;
                    case "<>":
                        return Math.Abs(left - right) >= 0.0000001;
                    default:
                        return Math.Abs(left - right) < 0.0000001;
                }
            }

            string text = value.Text ?? string.Empty;
            int comparison = string.Compare(text, criteria.Text, StringComparison.OrdinalIgnoreCase);
            switch (criteria.Operator) {
                case ">":
                    return comparison > 0;
                case "<":
                    return comparison < 0;
                case ">=":
                    return comparison >= 0;
                case "<=":
                    return comparison <= 0;
                case "<>":
                    return !MatchesTextCriteria(text, criteria.Text);
                default:
                    return MatchesTextCriteria(text, criteria.Text);
            }
        }

        private static bool MatchesTextCriteria(string text, string criteria) {
            if (criteria.IndexOf('*') < 0 && criteria.IndexOf('?') < 0) {
                return string.Equals(text, criteria, StringComparison.OrdinalIgnoreCase);
            }

            string pattern = "^" + Regex.Escape(criteria).Replace(@"\*", ".*").Replace(@"\?", ".") + "$";
            return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase, FormulaRegexTimeout);
        }

        private static bool TryGetSupportedDecimalPlaces(double value, out int digits) {
            digits = (int)Math.Round(value, MidpointRounding.AwayFromZero);
            return Math.Abs(value - digits) < 0.0000001 && digits >= 0 && digits <= 15;
        }

        private static double CalculateMedian(List<double> numbers) {
            var sorted = numbers.OrderBy(value => value).ToList();
            int middle = sorted.Count / 2;
            return sorted.Count % 2 == 1
                ? sorted[middle]
                : (sorted[middle - 1] + sorted[middle]) / 2d;
        }

        private static bool TryGetWholeNumber(double value, out int number) {
            number = (int)Math.Round(value, MidpointRounding.AwayFromZero);
            return Math.Abs(value - number) < 0.0000001;
        }

        private bool TryGetWholeNumberArgument(string token, out int number) {
            number = 0;
            return TryEvaluateFormulaOrNumeric(token, out double value)
                && TryGetWholeNumber(value, out number);
        }

        private static bool TryGetDateFromSerial(double value, out DateTime date) {
            date = default;
            try {
                date = DateTime.FromOADate(value).Date;
                return true;
            } catch (ArgumentException) {
                return false;
            }
        }

        private static string GetUnsupportedFormulaReason(string formula) {
            if (string.IsNullOrWhiteSpace(formula)) {
                return "Formula is empty.";
            }

            if (formula.Length > MaxSupportedFormulaLength) {
                return $"Formula is longer than {MaxSupportedFormulaLength} characters.";
            }

            try {
                if (formula.IndexOf(';') >= 0) {
                    return "Formula uses semicolon argument separators; OfficeIMO's lightweight evaluator expects Open XML comma-separated formulas.";
                }

                if (formula.IndexOf('&') >= 0) {
                    return "Formula uses the text concatenation operator, which OfficeIMO's lightweight evaluator does not currently support.";
                }

                if (formula.IndexOf('{') >= 0 || formula.IndexOf('}') >= 0) {
                    return "Formula uses array constants, which OfficeIMO's lightweight evaluator does not currently support.";
                }

                Match supportedFunctionMatch = SimpleFunctionFormulaRegex.Match(formula);
                if (supportedFunctionMatch.Success) {
                    string function = supportedFunctionMatch.Groups[1].Value.ToUpperInvariant();
                    return $"Formula uses supported function '{function}' with arguments OfficeIMO's lightweight evaluator cannot currently evaluate.";
                }

                Match functionMatch = FunctionNameFormulaRegex.Match(formula);
                if (functionMatch.Success) {
                    string function = functionMatch.Groups[1].Value.ToUpperInvariant();
                    return $"Function '{function}' is not supported by OfficeIMO's lightweight evaluator.";
                }
            } catch (RegexMatchTimeoutException) {
                return "Formula diagnostics timed out while parsing the formula.";
            }

            return "Formula is outside OfficeIMO's lightweight evaluator support.";
        }

        private bool TryResolveTextArgumentValues(IEnumerable<string> tokens, out List<string> values) {
            values = new List<string>();
            foreach (string token in tokens) {
                if (TryResolveFormulaRange(token, out var rangeValues)) {
                    foreach (var rangeValue in rangeValues) {
                        if (rangeValue.IsUnresolvedFormula) {
                            values.Clear();
                            return false;
                        }

                        values.Add(FormulaValueToText(rangeValue));
                    }

                    continue;
                }

                if (!TryResolveTextArgument(token, out string value)) {
                    values.Clear();
                    return false;
                }

                values.Add(value);
            }

            return true;
        }

        private bool TryResolveTextArgument(string token, out string text) {
            text = string.Empty;
            if (!TryResolveFormulaArgument(token, out FormulaArgumentValue value)) {
                return false;
            }

            if (value.IsUnresolvedFormula) {
                return false;
            }

            text = FormulaValueToText(value);
            return true;
        }

        private bool TryResolveBooleanArgument(string token, out bool value) {
            string trimmed = token.Trim();
            if (trimmed.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) {
                value = true;
                return true;
            }

            if (trimmed.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) {
                value = false;
                return true;
            }

            if (TryEvaluateFormulaOrNumeric(trimmed, out double numeric)) {
                value = Math.Abs(numeric) >= double.Epsilon;
                return true;
            }

            value = false;
            return false;
        }

        private static string FormulaValueToText(FormulaArgumentValue value) {
            return value.Text ?? value.Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private bool TryResolveFormulaArgument(string token, out FormulaArgumentValue value) {
            string trimmed = token.Trim();
            if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"') {
                value = new FormulaArgumentValue(null, trimmed.Substring(1, trimmed.Length - 2).Replace("\"\"", "\""));
                return true;
            }

            if (TryParseQualifiedFormulaCellReference(trimmed, out ExcelSheet sheet, out int row, out int column)) {
                value = sheet.ResolveCellArgument(row, column);
                return true;
            }

            if (TryResolveFormulaRangeReference(trimmed, out ExcelSheet rangeSheet, out int r1, out int c1, out int r2, out int c2)
                && r1 == r2
                && c1 == c2) {
                value = rangeSheet.ResolveCellArgument(r1, c1);
                return true;
            }

            if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric)) {
                value = new FormulaArgumentValue(numeric, trimmed);
                return true;
            }

            if (TryEvaluateFormulaValue(trimmed, out value)) {
                return true;
            }

            value = default;
            return false;
        }

        private static bool IsExactLookupMode(string token) {
            string value = token.Trim();
            return value == "0" || value.Equals("FALSE", StringComparison.OrdinalIgnoreCase);
        }

        private static bool FormulaValuesEqual(FormulaArgumentValue left, FormulaArgumentValue right) {
            if (left.Number.HasValue && right.Number.HasValue) {
                return Math.Abs(left.Number.Value - right.Number.Value) < 0.0000001;
            }

            string leftText = left.Text ?? left.Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
            string rightText = right.Text ?? right.Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
            return string.Equals(leftText, rightText, StringComparison.OrdinalIgnoreCase);
        }

        private bool TryResolveFormulaArguments(string args, out List<FormulaArgumentValue> values) {
            values = new List<FormulaArgumentValue>();
            foreach (string trimmed in SplitFormulaArguments(args)) {
                if (TryResolveFormulaRange(trimmed, out var rangeValues)) {
                    values.AddRange(rangeValues);
                    continue;
                }

                if (TryParseQualifiedFormulaCellReference(trimmed, out ExcelSheet sheetReference, out int cellRow, out int cellColumn)) {
                    values.Add(sheetReference.ResolveCellArgument(cellRow, cellColumn));
                    continue;
                }

                if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric)) {
                    values.Add(new FormulaArgumentValue(numeric, trimmed));
                    continue;
                }

                if (TryEvaluateFormulaValue(trimmed, out var formulaValue)) {
                    values.Add(formulaValue);
                    continue;
                }

                values.Clear();
                return false;
            }

            return true;
        }

        private static IReadOnlyList<string> SplitFormulaArguments(string args) {
            var tokens = new List<string>();
            var builder = new StringBuilder();
            int depth = 0;
            int bracketDepth = 0;
            bool inString = false;

            for (int index = 0; index < args.Length; index++) {
                char ch = args[index];
                if (ch == '"') {
                    builder.Append(ch);
                    if (inString && index + 1 < args.Length && args[index + 1] == '"') {
                        index++;
                        builder.Append(args[index]);
                        continue;
                    }

                    inString = !inString;
                    continue;
                }

                if (!inString && ch == '(') {
                    depth++;
                    builder.Append(ch);
                    continue;
                }

                if (!inString && ch == ')') {
                    depth--;
                    if (depth < 0) {
                        return Array.Empty<string>();
                    }

                    builder.Append(ch);
                    continue;
                }

                if (!inString && ch == '[') {
                    bracketDepth++;
                    builder.Append(ch);
                    continue;
                }

                if (!inString && ch == ']') {
                    bracketDepth--;
                    if (bracketDepth < 0) {
                        return Array.Empty<string>();
                    }

                    builder.Append(ch);
                    continue;
                }

                if (!inString && ch == ',' && depth == 0 && bracketDepth == 0) {
                    AddToken(tokens, builder);
                    continue;
                }

                builder.Append(ch);
            }

            if (depth != 0 || bracketDepth != 0 || inString) {
                return Array.Empty<string>();
            }

            AddToken(tokens, builder);
            return tokens;
        }

        private static void AddToken(List<string> tokens, StringBuilder builder) {
            string token = builder.ToString().Trim();
            if (token.Length > 0) {
                tokens.Add(token);
            }

            builder.Clear();
        }

        private bool TryResolveNumericOperand(string token, out double value) {
            token = token.Trim().Replace("$", string.Empty);
            if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) {
                return true;
            }

            if (!TryParseQualifiedFormulaCellReference(token, out ExcelSheet sheet, out int row, out int column)
                && (!TryResolveFormulaRangeReference(token, out sheet, out row, out column, out int r2, out int c2)
                    || row != r2
                    || column != c2)) {
                return false;
            }

            var argument = sheet.ResolveCellArgument(row, column);
            if (argument.Number.HasValue) {
                value = argument.Number.Value;
                return true;
            }

            return false;
        }

        private bool TryParseQualifiedFormulaCellReference(string token, out ExcelSheet sheet, out int row, out int column) {
            return TryParseQualifiedFormulaCellReference(token, null, out sheet, out row, out column);
        }

        private bool TryParseQualifiedFormulaCellReference(string token, ExcelSheet? defaultSheet, out ExcelSheet sheet, out int row, out int column) {
            sheet = this;
            row = 0;
            column = 0;

            if (!TrySplitQualifiedReference(token, out string? sheetName, out string reference)) {
                return false;
            }

            if (sheetName != null) {
                if (!TryGetFormulaReferenceSheet(sheetName, out sheet)) {
                    return false;
                }
            } else if (defaultSheet != null) {
                sheet = defaultSheet;
            }

            var cellRef = A1.ParseCellRef(reference.Replace("$", string.Empty));
            row = cellRef.Row;
            column = cellRef.Col;
            return row > 0
                && column > 0
                && row <= A1.MaxRows
                && column <= A1.MaxColumns;
        }

        private bool TryParseQualifiedFormulaRange(string token, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2) {
            return TryParseQualifiedFormulaRange(token, null, out sheet, out r1, out c1, out r2, out c2);
        }

        private bool TryParseQualifiedFormulaRange(string token, ExcelSheet? defaultSheet, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2) {
            sheet = this;
            r1 = 0;
            c1 = 0;
            r2 = 0;
            c2 = 0;

            if (!TrySplitQualifiedReference(token, out string? sheetName, out string reference)) {
                return false;
            }

            if (sheetName != null) {
                if (!TryGetFormulaReferenceSheet(sheetName, out sheet)) {
                    return false;
                }
            } else if (defaultSheet != null) {
                sheet = defaultSheet;
            }

            return A1.TryParseRange(reference.Replace("$", string.Empty), out r1, out c1, out r2, out c2);
        }

        private bool TryResolveFormulaRangeReference(string token, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2) {
            if (TryParseQualifiedFormulaRange(token, out sheet, out r1, out c1, out r2, out c2)) {
                return true;
            }

            if (TryParseQualifiedFormulaCellReference(token, out sheet, out r1, out c1)) {
                r2 = r1;
                c2 = c1;
                return true;
            }

            if (TryResolveTableReferenceRange(token, out sheet, out r1, out c1, out r2, out c2)) {
                return true;
            }

            return TryResolveDefinedNameRange(token, out sheet, out r1, out c1, out r2, out c2);
        }

        private bool TryResolveTableReferenceRange(string token, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2) {
            sheet = this;
            r1 = 0;
            c1 = 0;
            r2 = 0;
            c2 = 0;

            if (!TryParseStructuredTableReference(token, out string tableName, out var sections)) {
                return false;
            }

            WorkbookPart? workbookPart = _spreadSheetDocument.WorkbookPart;
            if (workbookPart == null) {
                return false;
            }

            foreach (var sheetElement in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (sheetElement.Id?.Value == null) {
                    continue;
                }

                if (workbookPart.GetPartById(sheetElement.Id.Value) is not WorksheetPart worksheetPart) {
                    continue;
                }

                foreach (var tablePart in worksheetPart.TableDefinitionParts) {
                    Table? table = tablePart.Table;
                    if (table == null
                        || (!string.Equals(table.Name?.Value, tableName, StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(table.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase))) {
                        continue;
                    }

                    sheet = string.Equals(Name, sheetElement.Name?.Value, StringComparison.OrdinalIgnoreCase)
                        ? this
                        : new ExcelSheet(_excelDocument, _spreadSheetDocument, sheetElement) {
                            _formulaEvaluationCache = _formulaEvaluationCache,
                            _formulaEvaluationStack = _formulaEvaluationStack
                        };
                    return TryResolveTableReferenceRange(table, sections, out r1, out c1, out r2, out c2);
                }
            }

            return false;
        }

        private static bool TryParseStructuredTableReference(string token, out string tableName, out List<string> sections) {
            string value = token.Trim();
            tableName = string.Empty;
            sections = new List<string>();
            if (value.Length == 0 || value.IndexOf('!') >= 0) {
                return false;
            }

            int bracketStart = value.IndexOf('[');
            tableName = bracketStart < 0 ? value : value.Substring(0, bracketStart);
            if (!IsFormulaDefinedNameToken(tableName)) {
                return false;
            }

            if (bracketStart < 0) {
                return true;
            }

            string specifier = value.Substring(bracketStart);
            if (specifier.Length >= 4 && specifier.StartsWith("[[", StringComparison.Ordinal) && specifier.EndsWith("]]", StringComparison.Ordinal)) {
                return TryParseStructuredTableSectionList(specifier.Substring(1, specifier.Length - 2), sections);
            }

            if (specifier.Length >= 2 && specifier[0] == '[' && specifier[specifier.Length - 1] == ']') {
                string section = specifier.Substring(1, specifier.Length - 2).Trim();
                if (section.Length == 0 || section.IndexOf('[') >= 0 || section.IndexOf(']') >= 0) {
                    return false;
                }

                sections.Add(section);
                return true;
            }

            return false;
        }

        private static bool TryParseStructuredTableSectionList(string value, List<string> sections) {
            int index = 0;
            while (index < value.Length) {
                if (value[index] != '[') {
                    return false;
                }

                int end = value.IndexOf(']', index + 1);
                if (end < 0) {
                    return false;
                }

                string section = value.Substring(index + 1, end - index - 1).Trim();
                if (section.Length == 0 || section.IndexOf('[') >= 0 || section.IndexOf(']') >= 0) {
                    return false;
                }

                sections.Add(section);
                index = end + 1;
                if (index == value.Length) {
                    return true;
                }

                if (value[index] != ',') {
                    return false;
                }

                index++;
            }

            return sections.Count > 0;
        }

        private static bool TryResolveTableReferenceRange(Table table, IReadOnlyList<string> sections, out int r1, out int c1, out int r2, out int c2) {
            r1 = 0;
            c1 = 0;
            r2 = 0;
            c2 = 0;

            if (table.Reference?.Value == null
                || !A1.TryParseRange(table.Reference.Value.Replace("$", string.Empty), out int tableR1, out int tableC1, out int tableR2, out int tableC2)) {
                return false;
            }

            uint headerRows = table.HeaderRowCount?.Value ?? 1U;
            uint totalsRows = table.TotalsRowShown?.Value == true
                ? Math.Max(1U, table.TotalsRowCount?.Value ?? 1U)
                : 0U;

            string? item = null;
            string area = "#Data";
            if (sections.Count == 1) {
                if (IsStructuredTableAreaSpecifier(sections[0])) {
                    area = sections[0];
                } else {
                    item = sections[0];
                }
            } else if (sections.Count == 2) {
                if (!IsStructuredTableAreaSpecifier(sections[0])) {
                    return false;
                }

                area = sections[0];
                item = sections[1];
            } else if (sections.Count != 0) {
                return false;
            }

            if (!TryResolveTableAreaRows(area, tableR1, tableR2, headerRows, totalsRows, out r1, out r2)) {
                return false;
            }

            c1 = tableC1;
            c2 = tableC2;
            if (!string.IsNullOrWhiteSpace(item)) {
                int offset = ResolveTableColumnOffset(table, item!);
                if (offset < 0) {
                    return false;
                }

                c1 = tableC1 + offset;
                c2 = c1;
            }

            return r1 <= r2 && c1 <= c2;
        }

        private static bool TryResolveTableAreaRows(string area, int tableR1, int tableR2, uint headerRows, uint totalsRows, out int r1, out int r2) {
            r1 = tableR1;
            r2 = tableR2;
            if (string.Equals(area, "#All", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (string.Equals(area, "#Headers", StringComparison.OrdinalIgnoreCase)) {
                if (headerRows == 0) {
                    return false;
                }

                r2 = tableR1 + (int)headerRows - 1;
                return r2 <= tableR2;
            }

            if (string.Equals(area, "#Totals", StringComparison.OrdinalIgnoreCase)) {
                if (totalsRows == 0) {
                    return false;
                }

                r1 = tableR2 - (int)totalsRows + 1;
                return r1 >= tableR1;
            }

            if (!string.Equals(area, "#Data", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            r1 = tableR1 + (int)headerRows;
            r2 = tableR2 - (int)totalsRows;
            return r1 <= r2;
        }

        private static bool IsStructuredTableAreaSpecifier(string section) {
            return string.Equals(section, "#All", StringComparison.OrdinalIgnoreCase)
                || string.Equals(section, "#Data", StringComparison.OrdinalIgnoreCase)
                || string.Equals(section, "#Headers", StringComparison.OrdinalIgnoreCase)
                || string.Equals(section, "#Totals", StringComparison.OrdinalIgnoreCase);
        }

        private static int ResolveTableColumnOffset(Table table, string columnName) {
            int index = 0;
            foreach (var tableColumn in table.TableColumns?.Elements<TableColumn>() ?? Enumerable.Empty<TableColumn>()) {
                if (string.Equals(tableColumn.Name?.Value, columnName, StringComparison.OrdinalIgnoreCase)) {
                    return index;
                }

                index++;
            }

            return -1;
        }

        private bool TryResolveDefinedNameRange(string token, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2) {
            sheet = this;
            r1 = 0;
            c1 = 0;
            r2 = 0;
            c2 = 0;

            if (!TrySplitQualifiedReference(token, out string? sheetName, out string name)
                || !IsFormulaDefinedNameToken(name)) {
                return false;
            }

            var definedNames = WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return false;
            }

            var sheets = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            int? localSheetIndex = null;
            ExcelSheet defaultSheet = this;
            if (sheetName != null) {
                int index = sheets.FindIndex(candidate => string.Equals(candidate.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
                if (index < 0 || !TryGetFormulaReferenceSheet(sheetName, out defaultSheet)) {
                    return false;
                }

                localSheetIndex = index;
            } else {
                int index = sheets.FindIndex(candidate => string.Equals(candidate.Name?.Value, Name, StringComparison.OrdinalIgnoreCase));
                if (index >= 0) {
                    localSheetIndex = index;
                }
            }

            DefinedName? definedName = null;
            if (localSheetIndex.HasValue) {
                definedName = definedNames.Elements<DefinedName>()
                    .FirstOrDefault(candidate => candidate.LocalSheetId?.Value == (uint)localSheetIndex.Value
                        && string.Equals(candidate.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
            }

            if (definedName == null && sheetName == null) {
                definedName = definedNames.Elements<DefinedName>()
                    .FirstOrDefault(candidate => candidate.LocalSheetId == null
                        && string.Equals(candidate.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
            }

            if (definedName == null || IsBuiltInFormulaDefinedName(definedName.Name?.Value)) {
                return false;
            }

            string reference = (definedName.Text ?? string.Empty).Trim();
            if (reference.StartsWith("=", StringComparison.Ordinal)) {
                reference = reference.Substring(1).Trim();
            }

            if (reference.Length == 0
                || reference.IndexOf(',') >= 0
                || reference.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0) {
                return false;
            }

            if (definedName.LocalSheetId?.Value is uint scopedIndex
                && scopedIndex < (uint)sheets.Count
                && sheets[(int)scopedIndex].Name?.Value is string scopedSheetName
                && !string.Equals(scopedSheetName, Name, StringComparison.OrdinalIgnoreCase)) {
                _ = TryGetFormulaReferenceSheet(scopedSheetName, out defaultSheet);
            }

            if (TryParseQualifiedFormulaRange(reference, defaultSheet, out sheet, out r1, out c1, out r2, out c2)) {
                return true;
            }

            if (TryParseQualifiedFormulaCellReference(reference, defaultSheet, out sheet, out r1, out c1)) {
                r2 = r1;
                c2 = c1;
                return true;
            }

            return false;
        }

        private static bool IsFormulaDefinedNameToken(string token) {
            if (string.IsNullOrWhiteSpace(token)
                || string.Equals(token, "TRUE", StringComparison.OrdinalIgnoreCase)
                || string.Equals(token, "FALSE", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            char first = token[0];
            if (!char.IsLetter(first) && first != '_') {
                return false;
            }

            foreach (char character in token) {
                if (!char.IsLetterOrDigit(character) && character != '_' && character != '.') {
                    return false;
                }
            }

            return true;
        }

        private static bool IsBuiltInFormulaDefinedName(string? name) {
            return !string.IsNullOrWhiteSpace(name)
                && name!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase);
        }

        private bool TryGetFormulaReferenceSheet(string sheetName, out ExcelSheet sheet) {
            sheet = this;
            if (string.Equals(Name, sheetName, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            var sheetElement = WorkbookRoot.Sheets?
                .Elements<Sheet>()
                .FirstOrDefault(candidate => string.Equals(candidate.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheetElement?.Id == null) {
                return false;
            }

            sheet = new ExcelSheet(_excelDocument, _spreadSheetDocument, sheetElement) {
                _formulaEvaluationCache = _formulaEvaluationCache,
                _formulaEvaluationStack = _formulaEvaluationStack
            };
            return true;
        }

        private static bool TrySplitQualifiedReference(string token, out string? sheetName, out string reference) {
            string value = token.Trim();
            sheetName = null;
            reference = value;
            if (value.Length == 0) {
                return false;
            }

            int separator = value.LastIndexOf('!');
            if (separator < 0) {
                return true;
            }

            if (separator == 0 || separator == value.Length - 1) {
                return false;
            }

            sheetName = NormalizeFormulaSheetName(value.Substring(0, separator));
            reference = value.Substring(separator + 1).Trim();
            return !string.IsNullOrWhiteSpace(sheetName) && reference.Length > 0;
        }

        private static string NormalizeFormulaSheetName(string token) {
            string value = token.Trim();
            if (value.Length >= 2 && value[0] == '\'' && value[value.Length - 1] == '\'') {
                value = value.Substring(1, value.Length - 2).Replace("''", "'");
            }

            return value;
        }

        private FormulaArgumentValue ResolveCellArgument(int row, int column) {
            var cell = TryGetExistingCell(row, column);
            bool unresolvedFormula = false;
            if (cell?.CellFormula != null && _formulaEvaluationCache != null) {
                if (TryEvaluateFormulaCellValue(cell, out FormulaArgumentValue formulaResult)) {
                    return formulaResult;
                }

                unresolvedFormula = true;
            }

            var value = GetCellValueSnapshot(row, column);
            if (unresolvedFormula && value.Value == null && string.IsNullOrEmpty(value.CachedText)) {
                return FormulaArgumentValue.UnresolvedFormula();
            }

            if (value.Value is double d) {
                return new FormulaArgumentValue(d, value.CachedText);
            }

            if (double.TryParse(value.CachedText, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) {
                return new FormulaArgumentValue(parsed, value.CachedText);
            }

            return new FormulaArgumentValue(null, value.Value?.ToString());
        }

        private static string? NormalizeFormulaCellReference(string? reference) {
            if (string.IsNullOrWhiteSpace(reference)) {
                return null;
            }

            var cellRef = A1.ParseCellRef(reference!.Trim().Replace("$", string.Empty));
            if (cellRef.Row <= 0
                || cellRef.Col <= 0
                || cellRef.Row > A1.MaxRows
                || cellRef.Col > A1.MaxColumns) {
                return null;
            }

            return A1.CellReference(cellRef.Row, cellRef.Col);
        }

        private string GetFormulaEvaluationCacheKey(string reference) {
            return Name + "!" + reference;
        }

        private readonly struct FormulaArgumentValue {
            internal FormulaArgumentValue(double? number, string? text, bool isUnresolvedFormula = false) {
                Number = number;
                Text = text;
                IsUnresolvedFormula = isUnresolvedFormula;
            }

            internal double? Number { get; }
            internal string? Text { get; }
            internal bool IsUnresolvedFormula { get; }
            internal bool HasValue => Number.HasValue || Text != null;

            internal static FormulaArgumentValue UnresolvedFormula() {
                return new FormulaArgumentValue(null, null, isUnresolvedFormula: true);
            }
        }

        private readonly struct FormulaCriteria {
            internal FormulaCriteria(string op, string text, double? number) {
                Operator = op;
                Text = text;
                Number = number;
            }

            internal string Operator { get; }
            internal string Text { get; }
            internal double? Number { get; }
        }
    }
}
