using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int MaxSupportedFormulaLength = 8192;
        private static readonly TimeSpan FormulaRegexTimeout = TimeSpan.FromMilliseconds(100);

        private static readonly Regex SimpleFunctionFormulaRegex = new Regex(
            @"^\s*=?\s*(SUM|AVERAGE|MIN|MAX|COUNT|COUNTA|COUNTIF|SUMIF|AVERAGEIF|COUNTIFS|SUMIFS|AVERAGEIFS|PRODUCT|MEDIAN|LARGE|SMALL|SUMPRODUCT|VLOOKUP|HLOOKUP|XLOOKUP|ABS|SIGN|ROUND|ROUNDUP|ROUNDDOWN|TRUNC|INT|CEILING|FLOOR|POWER|SQRT|LN|LOG10|EXP|PI|RADIANS|DEGREES|MOD|DATE|TIME|TODAY|NOW|YEAR|MONTH|DAY|HOUR|MINUTE|SECOND|EDATE|EOMONTH|DAYS|WEEKDAY|NETWORKDAYS|IF|AND|OR|NOT|IFERROR)\s*\((.*)\)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        private static readonly Regex SimpleBinaryFormulaRegex = new Regex(
            @"^\s*=?\s*([A-Z]+[0-9]+|-?\d+(?:\.\d+)?)\s*([+\-*/])\s*([A-Z]+[0-9]+|-?\d+(?:\.\d+)?)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        private static readonly Regex SimpleComparisonFormulaRegex = new Regex(
            @"^\s*([A-Z]+[0-9]+|-?\d+(?:\.\d+)?)\s*(>=|<=|<>|=|>|<)\s*([A-Z]+[0-9]+|-?\d+(?:\.\d+)?)\s*$",
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
                    bool supported = TryEvaluateFormula(formula, out _);
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

                    if (!TryResolveFormulaArguments(args, out var values)) {
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

        private bool TryEvaluateLookupFunction(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (function == "XLOOKUP") {
                return TryEvaluateXLookup(tokens, out result);
            }

            if (tokens.Count != 4
                || !TryResolveFormulaArgument(tokens[0], out var lookupValue)
                || !TryGetWholeNumberArgument(tokens[2], out int resultIndex)
                || !IsExactLookupMode(tokens[3])
                || !A1.TryParseRange(tokens[1].Trim().Replace("$", string.Empty), out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }

            if (function == "VLOOKUP") {
                int width = c2 - c1 + 1;
                if (resultIndex < 1 || resultIndex > width) {
                    return false;
                }

                int resultColumn = c1 + resultIndex - 1;
                for (int row = r1; row <= r2; row++) {
                    if (!FormulaValuesEqual(ResolveCellArgument(row, c1), lookupValue)) {
                        continue;
                    }

                    var value = ResolveCellArgument(row, resultColumn);
                    if (!value.Number.HasValue) {
                        return false;
                    }

                    result = value.Number.Value;
                    return true;
                }

                return false;
            }

            int height = r2 - r1 + 1;
            if (resultIndex < 1 || resultIndex > height) {
                return false;
            }

            int resultRow = r1 + resultIndex - 1;
            for (int column = c1; column <= c2; column++) {
                if (!FormulaValuesEqual(ResolveCellArgument(r1, column), lookupValue)) {
                    continue;
                }

                var value = ResolveCellArgument(resultRow, column);
                if (!value.Number.HasValue) {
                    return false;
                }

                result = value.Number.Value;
                return true;
            }

            return false;
        }

        private bool TryEvaluateXLookup(IReadOnlyList<string> tokens, out double result) {
            result = 0;
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
                if (!returnValue.Number.HasValue) {
                    return false;
                }

                result = returnValue.Number.Value;
                return true;
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
            if (token.IndexOf(':') >= 0) {
                if (!TryResolveFormulaRange(token, out var values)) {
                    return false;
                }

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
            if (!A1.TryParseRange(token.Trim().Replace("$", string.Empty), out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }

            for (int row = r1; row <= r2; row++) {
                for (int column = c1; column <= c2; column++) {
                    values.Add(ResolveCellArgument(row, column));
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

            return "Formula is outside OfficeIMO's lightweight evaluator support.";
        }

        private bool TryResolveFormulaArgument(string token, out FormulaArgumentValue value) {
            string trimmed = token.Trim();
            if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"') {
                value = new FormulaArgumentValue(null, trimmed.Substring(1, trimmed.Length - 2));
                return true;
            }

            if (TryParseFormulaCellReference(trimmed, out var cellRef)) {
                value = ResolveCellArgument(cellRef.Row, cellRef.Col);
                return true;
            }

            if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric)
                || TryEvaluateFormula(trimmed, out numeric)) {
                value = new FormulaArgumentValue(numeric, trimmed);
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

                if (TryEvaluateFormula(trimmed, out numeric)) {
                    values.Add(new FormulaArgumentValue(numeric, trimmed));
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

            foreach (char ch in args) {
                if (ch == '(') {
                    depth++;
                    builder.Append(ch);
                    continue;
                }

                if (ch == ')') {
                    depth--;
                    if (depth < 0) {
                        return Array.Empty<string>();
                    }

                    builder.Append(ch);
                    continue;
                }

                if (ch == ',' && depth == 0) {
                    AddToken(tokens, builder);
                    continue;
                }

                builder.Append(ch);
            }

            if (depth != 0) {
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
