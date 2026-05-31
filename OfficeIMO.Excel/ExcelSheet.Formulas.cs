using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int MaxSupportedFormulaLength = 8192;
        private static readonly TimeSpan FormulaRegexTimeout = TimeSpan.FromSeconds(1);
        private Dictionary<string, FormulaArgumentValue>? _formulaEvaluationCache;
        private HashSet<string>? _formulaEvaluationStack;

        private static readonly Regex SimpleFunctionFormulaRegex = new Regex(
            @"^\s*=?\s*(SUM|AVERAGE|AVERAGEA|MIN|MINA|MAX|MAXA|COUNT|COUNTA|COUNTBLANK|SUBTOTAL|COUNTIF|SUMIF|AVERAGEIF|COUNTIFS|SUMIFS|AVERAGEIFS|MINIFS|MAXIFS|PRODUCT|MEDIAN|LARGE|SMALL|MODE\.SNGL|MODE|GEOMEAN|HARMEAN|AVEDEV|DEVSQ|SUMXMY2|SUMX2MY2|SUMX2PY2|SUMSQ|SUMPRODUCT|STDEV\.S|STDEV\.P|VAR\.S|VAR\.P|PERCENTILE\.INC|PERCENTILE\.EXC|QUARTILE\.INC|QUARTILE\.EXC|PERCENTRANK\.INC|PERCENTRANK\.EXC|RANK\.EQ|RANK\.AVG|COVAR|COVARIANCE\.P|COVARIANCE\.S|CORREL|SLOPE|INTERCEPT|RSQ|FORECAST\.LINEAR|PMT|PV|FV|NPER|NPV|VLOOKUP|HLOOKUP|XLOOKUP|INDEX|MATCH|XMATCH|ABS|SIGN|ROUND|ROUNDUP|ROUNDDOWN|MROUND|TRUNC|INT|CEILING\.MATH|FLOOR\.MATH|CEILING|FLOOR|POWER|SQRT|LN|LOG10|EXP|PI|RADIANS|DEGREES|MOD|ROW|COLUMN|ROWS|COLUMNS|DATE|TIME|DATEVALUE|TIMEVALUE|TODAY|NOW|YEAR|MONTH|DAY|HOUR|MINUTE|SECOND|DATEDIF|YEARFRAC|EDATE|EOMONTH|DAYS|DAYS360|WEEKDAY|WEEKNUM|ISOWEEKNUM|NETWORKDAYS|WORKDAY\.INTL|WORKDAY|IF|IFS|SWITCH|CHOOSE|ISBLANK|ISNUMBER|ISTEXT|ISERROR|ISERR|ISNA|ISFORMULA|AND|OR|NOT|IFERROR|IFNA|CONCAT|CONCATENATE|TEXT|TEXTJOIN|TEXTBEFORE|TEXTAFTER|FORMULATEXT|LEFT|RIGHT|MID|LEN|TRIM|UPPER|LOWER|PROPER|SUBSTITUTE|FIND|SEARCH|VALUE|EXACT|REPT)\s*\((.*)\)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        private static readonly Regex FunctionNameFormulaRegex = new Regex(
            @"^\s*=?\s*([A-Za-z][A-Za-z0-9_.]*)\s*\(",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        private static readonly Regex FormulaReferenceRegex = new Regex(
            @"(?<![A-Za-z0-9_\.])(?<reference>(?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_ .]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)",
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
                MaterializePendingDirectCellValues();

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

            if (result.IsError) {
                cell.CellValue = new CellValue(result.ErrorCode ?? "#VALUE!");
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Error;
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
                    IReadOnlyList<string> dependencies = GetFormulaDependencies(formula);
                    IReadOnlyList<string> dependencyIssues = GetFormulaDependencyIssues(formula, cell.CellReference?.Value, dependencies);
                    formulas.Add(new ExcelFormulaCellInfo(
                        Name,
                        cell.CellReference?.Value ?? string.Empty,
                        formula,
                        cell.CellValue?.Text,
                        cell.CellFormula.CalculateCell?.Value ?? false,
                        supported,
                        supported ? null : GetUnsupportedFormulaReason(formula),
                        dependencies,
                        dependencyIssues));
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
                    if (function == "IFERROR" && TryEvaluateIfErrorValue(args, out result)) {
                        return true;
                    }

                    if (function == "IFNA" && TryEvaluateIfNaValue(args, out result)) {
                        return true;
                    }

                    if (function == "IF" && TryEvaluateIfValue(args, out result)) {
                        return true;
                    }

                    if (function == "IFS" && TryEvaluateIfsValue(args, out result)) {
                        return true;
                    }

                    if (function == "SWITCH" && TryEvaluateSwitchValue(args, out result)) {
                        return true;
                    }

                    if (function == "CHOOSE" && TryEvaluateChooseValue(args, out result)) {
                        return true;
                    }

                    if ((function == "ISBLANK" || function == "ISNUMBER" || function == "ISTEXT" || function == "ISERROR" || function == "ISERR" || function == "ISNA" || function == "ISFORMULA")
                        && TryEvaluateInfoFunction(function, args, out result)) {
                        return true;
                    }

                    if (TryEvaluateTextFunction(function, args, out result)) {
                        return true;
                    }

                    if ((function == "VLOOKUP" || function == "HLOOKUP" || function == "XLOOKUP")
                        && TryEvaluateLookupValue(function, args, out result)) {
                        return true;
                    }

                    if (function == "INDEX" && TryEvaluateIndexValue(args, out result)) {
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
                    if (function == "IFERROR" || function == "IFNA") {
                        if (!TryEvaluateErrorFallback(function, args, out result)) {
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

                    if (function == "IFS") {
                        if (!TryEvaluateIfsValue(args, out FormulaArgumentValue ifsResult) || !ifsResult.Number.HasValue) {
                            return false;
                        }

                        result = ifsResult.Number.Value;
                        return true;
                    }

                    if (function == "SWITCH") {
                        if (!TryEvaluateSwitchValue(args, out FormulaArgumentValue switchResult) || !switchResult.Number.HasValue) {
                            return false;
                        }

                        result = switchResult.Number.Value;
                        return true;
                    }

                    if (function == "CHOOSE") {
                        if (!TryEvaluateChooseValue(args, out FormulaArgumentValue chooseResult) || !chooseResult.Number.HasValue) {
                            return false;
                        }

                        result = chooseResult.Number.Value;
                        return true;
                    }

                    if (function == "ISBLANK" || function == "ISNUMBER" || function == "ISTEXT" || function == "ISERROR" || function == "ISERR" || function == "ISNA" || function == "ISFORMULA") {
                        if (!TryEvaluateInfoFunction(function, args, out FormulaArgumentValue infoResult) || !infoResult.Number.HasValue) {
                            return false;
                        }

                        result = infoResult.Number.Value;
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

                    if (function == "COUNTBLANK") {
                        return TryEvaluateCountBlank(args, out result);
                    }

                    if (function == "ROW" || function == "COLUMN" || function == "ROWS" || function == "COLUMNS") {
                        return TryEvaluateReferenceShapeFunction(function, args, out result);
                    }

                    if (function == "SUBTOTAL") {
                        return TryEvaluateSubtotal(args, out result);
                    }

                    if (function == "COUNTIF" || function == "SUMIF" || function == "AVERAGEIF") {
                        return TryEvaluateConditionalAggregate(function, args, out result);
                    }

                    if (function == "COUNTIFS" || function == "SUMIFS" || function == "AVERAGEIFS" || function == "MINIFS" || function == "MAXIFS") {
                        return TryEvaluateMultiCriteriaAggregate(function, args, out result);
                    }

                    if (function == "DATE" || function == "TIME" || function == "DATEVALUE" || function == "TIMEVALUE" || function == "TODAY" || function == "NOW"
                        || function == "YEAR" || function == "MONTH" || function == "DAY"
                        || function == "HOUR" || function == "MINUTE" || function == "SECOND"
                        || function == "DATEDIF" || function == "YEARFRAC"
                        || function == "EDATE" || function == "EOMONTH" || function == "DAYS" || function == "DAYS360"
                        || function == "WEEKDAY" || function == "WEEKNUM" || function == "ISOWEEKNUM" || function == "NETWORKDAYS"
                        || function == "WORKDAY" || function == "WORKDAY.INTL") {
                        return TryEvaluateDateTimeFunction(function, args, out result);
                    }

                    if (function == "SUMPRODUCT") {
                        return TryEvaluateSumProduct(args, out result);
                    }

                    if (function == "STDEV.S" || function == "STDEV.P" || function == "VAR.S" || function == "VAR.P"
                        || function == "MODE.SNGL" || function == "MODE" || function == "GEOMEAN" || function == "HARMEAN"
                        || function == "AVEDEV" || function == "DEVSQ"
                        || function == "SUMXMY2" || function == "SUMX2MY2" || function == "SUMX2PY2"
                        || function == "PERCENTILE.INC" || function == "PERCENTILE.EXC"
                        || function == "QUARTILE.INC" || function == "QUARTILE.EXC"
                        || function == "PERCENTRANK.INC" || function == "PERCENTRANK.EXC"
                        || function == "RANK.EQ" || function == "RANK.AVG"
                        || function == "COVAR" || function == "COVARIANCE.P" || function == "COVARIANCE.S"
                        || function == "CORREL" || function == "SLOPE" || function == "INTERCEPT" || function == "RSQ"
                        || function == "FORECAST.LINEAR") {
                        return TryEvaluateStatisticalFunction(function, args, out result);
                    }

                    if (function == "PMT" || function == "PV" || function == "FV" || function == "NPER" || function == "NPV") {
                        return TryEvaluateFinancialFunction(function, args, out result);
                    }

                    if (function == "VLOOKUP" || function == "HLOOKUP" || function == "XLOOKUP") {
                        return TryEvaluateLookupFunction(function, args, out result);
                    }

                    if (function == "INDEX") {
                        if (!TryEvaluateIndexValue(args, out FormulaArgumentValue indexValue) || !indexValue.Number.HasValue) {
                            return false;
                        }

                        result = indexValue.Number.Value;
                        return true;
                    }

                    if (function == "MATCH" || function == "XMATCH") {
                        return TryEvaluateMatchFunction(function, args, out result);
                    }

                    if (TryEvaluateTextFunction(function, args, out FormulaArgumentValue textFunctionResult)
                        && textFunctionResult.Number.HasValue) {
                        result = textFunctionResult.Number.Value;
                        return true;
                    }

                    if (!TryResolveFormulaArguments(args, out var values) || values.Any(value => value.IsUnresolvedFormula)) {
                        return false;
                    }

                    if (function == "COUNTA") {
                        result = values.Count(v => v.HasValue || !string.IsNullOrEmpty(v.Text));
                        return true;
                    }

                    var numbers = values.Where(v => v.Number.HasValue).Select(v => v.Number!.Value).ToList();
                    if (function == "AVERAGEA" || function == "MINA" || function == "MAXA") {
                        if (!TryConvertFormulaAValues(values, out var aValues) || aValues.Count == 0) {
                            return false;
                        }

                        if (function == "AVERAGEA") {
                            result = aValues.Average();
                            return IsFinite(result);
                        }

                        result = function == "MINA" ? aValues.Min() : aValues.Max();
                        return IsFinite(result);
                    }

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
                        if (numbers.Count != 2 || !TryGetSupportedDecimalPlaces(numbers[1], out int digits)) {
                            return false;
                        }

                        result = RoundAtDigits(numbers[0], digits, MidpointRounding.AwayFromZero);
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

                    if (function == "MROUND") {
                        if (numbers.Count != 2 || !TryEvaluateMRound(numbers[0], numbers[1], out result)) {
                            return false;
                        }

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

                    if (function == "CEILING.MATH" || function == "FLOOR.MATH") {
                        if (numbers.Count < 1 || numbers.Count > 3 || !TryEvaluateMathRoundFunction(function, numbers, out result)) {
                            return false;
                        }

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
                    else if (function == "SUMSQ") result = numbers.Sum(value => value * value);
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
            if (function == "FORMULATEXT") {
                if (tokens.Count != 1 || !TryGetFormulaTextArgument(tokens[0], out string formulaText)) {
                    return false;
                }

                result = new FormulaArgumentValue(null, formulaText.StartsWith("=", StringComparison.Ordinal) ? formulaText : "=" + formulaText);
                return true;
            }

            if (function == "CONCAT" || function == "CONCATENATE") {
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

            if (function == "TEXT") {
                if (tokens.Count != 2
                    || !TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue value)
                    || value.IsUnresolvedFormula
                    || !value.HasValue
                    || !TryResolveTextArgument(tokens[1], out string format)
                    || !TryFormatTextFunctionValue(value, format, out string formatted)) {
                    return false;
                }

                result = new FormulaArgumentValue(null, formatted);
                return true;
            }

            if (function == "TEXTBEFORE" || function == "TEXTAFTER") {
                return TryEvaluateTextBeforeAfterFunction(function == "TEXTBEFORE", tokens, out result);
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

            if (function == "UPPER" || function == "LOWER" || function == "PROPER") {
                if (tokens.Count != 1 || !TryResolveTextArgument(tokens[0], out string text)) {
                    return false;
                }

                string transformed = function == "UPPER"
                    ? text.ToUpperInvariant()
                    : function == "LOWER"
                        ? text.ToLowerInvariant()
                        : ToProperCase(text);
                result = new FormulaArgumentValue(null, transformed);
                return true;
            }

            if (function == "SUBSTITUTE") {
                if (tokens.Count < 3
                    || tokens.Count > 4
                    || !TryResolveTextArgument(tokens[0], out string text)
                    || !TryResolveTextArgument(tokens[1], out string oldText)
                    || !TryResolveTextArgument(tokens[2], out string newText)) {
                    return false;
                }

                if (tokens.Count == 4) {
                    if (!TryGetWholeNumberArgument(tokens[3], out int occurrence) || occurrence < 1) {
                        return false;
                    }

                    result = new FormulaArgumentValue(null, SubstituteTextOccurrence(text, oldText, newText, occurrence));
                    return true;
                }

                result = new FormulaArgumentValue(null, oldText.Length == 0 ? text : text.Replace(oldText, newText));
                return true;
            }

            if (function == "FIND" || function == "SEARCH") {
                if (tokens.Count < 2
                    || tokens.Count > 3
                    || !TryResolveTextArgument(tokens[0], out string findText)
                    || !TryResolveTextArgument(tokens[1], out string withinText)) {
                    return false;
                }

                int start = 1;
                if (tokens.Count == 3 && !TryGetWholeNumberArgument(tokens[2], out start)) {
                    return false;
                }

                if (start < 1 || start > withinText.Length + 1) {
                    return false;
                }

                StringComparison comparison = function == "SEARCH"
                    ? StringComparison.OrdinalIgnoreCase
                    : StringComparison.Ordinal;
                int foundIndex = withinText.IndexOf(findText, start - 1, comparison);
                if (foundIndex < 0) {
                    return false;
                }

                double position = foundIndex + 1;
                result = new FormulaArgumentValue(position, position.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (function == "VALUE") {
                if (tokens.Count != 1
                    || !TryResolveTextArgument(tokens[0], out string text)
                    || !TryParseValueFunctionNumber(text, out double number)) {
                    return false;
                }

                result = new FormulaArgumentValue(number, number.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (function == "EXACT") {
                if (tokens.Count != 2
                    || !TryResolveTextArgument(tokens[0], out string left)
                    || !TryResolveTextArgument(tokens[1], out string right)) {
                    return false;
                }

                double value = string.Equals(left, right, StringComparison.Ordinal) ? 1d : 0d;
                result = new FormulaArgumentValue(value, value.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (function == "REPT") {
                if (tokens.Count != 2
                    || !TryResolveTextArgument(tokens[0], out string text)
                    || !TryGetWholeNumberArgument(tokens[1], out int count)
                    || count < 0) {
                    return false;
                }

                if (text.Length > 0 && count > MaxSupportedFormulaLength / text.Length) {
                    return false;
                }

                result = new FormulaArgumentValue(null, string.Concat(Enumerable.Repeat(text, count)));
                return true;
            }

            return false;
        }

        private bool TryEvaluateTextBeforeAfterFunction(bool before, IReadOnlyList<string> tokens, out FormulaArgumentValue result) {
            result = default;
            if (tokens.Count < 2
                || tokens.Count > 6
                || !TryResolveTextArgument(tokens[0], out string text)
                || !TryResolveTextArgument(tokens[1], out string delimiter)
                || delimiter.Length == 0) {
                return false;
            }

            int instance = 1;
            if (tokens.Count >= 3 && !TryGetWholeNumberArgument(tokens[2], out instance)) {
                return false;
            }

            if (instance == 0) {
                return false;
            }

            int matchMode = 0;
            if (tokens.Count >= 4 && (!TryGetWholeNumberArgument(tokens[3], out matchMode) || (matchMode != 0 && matchMode != 1))) {
                return false;
            }

            bool matchEnd = false;
            if (tokens.Count >= 5 && !TryResolveBooleanArgument(tokens[4], out matchEnd)) {
                return false;
            }

            string? ifNotFound = null;
            if (tokens.Count >= 6 && !TryResolveTextArgument(tokens[5], out ifNotFound)) {
                return false;
            }

            StringComparison comparison = matchMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            if (!TryFindTextDelimiterOccurrence(text, delimiter, instance, comparison, out int index)) {
                if (matchEnd) {
                    result = new FormulaArgumentValue(null, before ? text : string.Empty);
                    return true;
                }

                if (ifNotFound != null) {
                    result = new FormulaArgumentValue(null, ifNotFound);
                    return true;
                }

                return false;
            }

            string extracted = before
                ? text.Substring(0, index)
                : text.Substring(index + delimiter.Length);
            result = new FormulaArgumentValue(null, extracted);
            return true;
        }

        private static bool TryFindTextDelimiterOccurrence(string text, string delimiter, int instance, StringComparison comparison, out int index) {
            if (instance > 0) {
                int searchStart = 0;
                for (int current = 1; current <= instance; current++) {
                    index = text.IndexOf(delimiter, searchStart, comparison);
                    if (index < 0) {
                        return false;
                    }

                    if (current == instance) {
                        return true;
                    }

                    searchStart = index + delimiter.Length;
                    if (searchStart > text.Length) {
                        return false;
                    }
                }
            } else {
                if (text.Length == 0) {
                    index = -1;
                    return false;
                }

                int searchStart = text.Length - 1;
                for (int current = -1; current >= instance; current--) {
                    index = text.LastIndexOf(delimiter, searchStart, comparison);
                    if (index < 0) {
                        return false;
                    }

                    if (current == instance) {
                        return true;
                    }

                    searchStart = index - 1;
                    if (searchStart < 0) {
                        return false;
                    }
                }
            }

            index = -1;
            return false;
        }

        private static string SubstituteTextOccurrence(string text, string oldText, string newText, int occurrence) {
            if (oldText.Length == 0) {
                return text;
            }

            int startIndex = 0;
            int currentOccurrence = 0;
            while (startIndex <= text.Length) {
                int index = text.IndexOf(oldText, startIndex, StringComparison.Ordinal);
                if (index < 0) {
                    return text;
                }

                currentOccurrence++;
                if (currentOccurrence == occurrence) {
                    return text.Substring(0, index) + newText + text.Substring(index + oldText.Length);
                }

                startIndex = index + oldText.Length;
            }

            return text;
        }

        private static bool TryParseValueFunctionNumber(string text, out double number) {
            string normalized = text.Trim();
            if (normalized.Length == 0) {
                number = 0d;
                return false;
            }

            normalized = normalized.Replace(",", string.Empty);
            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out number);
        }

        private static string ToProperCase(string text) {
            var builder = new StringBuilder(text.Length);
            bool capitalizeNext = true;
            foreach (char character in text) {
                if (char.IsLetter(character)) {
                    builder.Append(capitalizeNext
                        ? char.ToUpperInvariant(character)
                        : char.ToLowerInvariant(character));
                    capitalizeNext = false;
                    continue;
                }

                builder.Append(character);
                capitalizeNext = true;
            }

            return builder.ToString();
        }

        private bool TryEvaluateIndexValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 2
                || tokens.Count > 3
                || !TryResolveFormulaRangeReference(tokens[0], out ExcelSheet rangeSheet, out int r1, out int c1, out int r2, out int c2)
                || !TryGetWholeNumberArgument(tokens[1], out int rowIndex)) {
                return false;
            }

            int rowCount = r2 - r1 + 1;
            int columnCount = c2 - c1 + 1;
            int columnIndex;
            if (tokens.Count == 3) {
                if (!TryGetWholeNumberArgument(tokens[2], out columnIndex)) {
                    return false;
                }
            } else if (columnCount == 1) {
                columnIndex = 1;
            } else if (rowCount == 1) {
                columnIndex = rowIndex;
                rowIndex = 1;
            } else {
                return false;
            }

            if (rowIndex < 1 || rowIndex > rowCount || columnIndex < 1 || columnIndex > columnCount) {
                return false;
            }

            result = rangeSheet.ResolveCellArgument(r1 + rowIndex - 1, c1 + columnIndex - 1);
            return result.HasValue;
        }

        private bool TryEvaluateMatchFunction(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            int maxTokens = function == "XMATCH" ? 4 : 3;
            if (tokens.Count < 2
                || tokens.Count > maxTokens
                || !TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue lookupValue)
                || !TryResolveFormulaRangeReference(tokens[1], out ExcelSheet rangeSheet, out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }

            bool vertical = c1 == c2;
            bool horizontal = r1 == r2;
            if (!vertical && !horizontal) {
                return false;
            }

            int matchMode = function == "XMATCH" ? 0 : 1;
            if (tokens.Count >= 3 && !TryGetWholeNumberArgument(tokens[2], out matchMode)) {
                return false;
            }

            if (function == "MATCH" && matchMode != -1 && matchMode != 0 && matchMode != 1) {
                return false;
            }

            if (function == "XMATCH" && matchMode != -1 && matchMode != 0 && matchMode != 1) {
                return false;
            }

            int searchMode = 1;
            if (function == "XMATCH"
                && tokens.Count >= 4
                && (!TryGetWholeNumberArgument(tokens[3], out searchMode) || (searchMode != 1 && searchMode != -1))) {
                return false;
            }

            int lookupMode = function == "MATCH" ? -matchMode : matchMode;
            if (!TryResolveFormulaRange(tokens[1], out var lookupValues)
                || !TryFindLookupPosition(lookupValue, lookupValues, lookupMode, searchMode, out int position)) {
                return false;
            }

            result = position;
            return true;
        }

        private static bool TryFindLookupPosition(FormulaArgumentValue lookupValue, IReadOnlyList<FormulaArgumentValue> lookupValues, int matchMode, int searchMode, out int position) {
            position = 0;
            if (lookupValues.Count == 0) {
                return false;
            }

            int start = searchMode == -1 ? lookupValues.Count - 1 : 0;
            int end = searchMode == -1 ? -1 : lookupValues.Count;
            int step = searchMode == -1 ? -1 : 1;

            for (int index = start; index != end; index += step) {
                if (!FormulaValuesEqual(lookupValues[index], lookupValue)) {
                    continue;
                }

                position = index + 1;
                return true;
            }

            if (matchMode == 0 || !lookupValue.Number.HasValue) {
                return false;
            }

            double bestDelta = double.MaxValue;
            int bestPosition = 0;
            for (int index = start; index != end; index += step) {
                FormulaArgumentValue candidate = lookupValues[index];
                if (!candidate.Number.HasValue) {
                    continue;
                }

                double delta = candidate.Number.Value - lookupValue.Number.Value;
                bool eligible = matchMode < 0
                    ? delta <= 0d
                    : delta >= 0d;
                if (!eligible) {
                    continue;
                }

                double distance = Math.Abs(delta);
                if (distance < bestDelta) {
                    bestDelta = distance;
                    bestPosition = index + 1;
                }
            }

            if (bestPosition == 0) {
                return false;
            }

            position = bestPosition;
            return true;
        }

        private static bool TryFormatTextFunctionValue(FormulaArgumentValue value, string format, out string formatted) {
            formatted = string.Empty;
            if (string.IsNullOrWhiteSpace(format)) {
                return false;
            }

            if (format == "@") {
                formatted = FormulaValueToText(value);
                return true;
            }

            if (LooksLikeDateTextFormat(format)) {
                if (!value.Number.HasValue || !TryGetDateTimeFromSerial(value.Number.Value, out DateTime date)) {
                    return false;
                }

                string dotNetFormat = ConvertExcelDateTextFormat(format);
                try {
                    formatted = date.ToString(dotNetFormat, CultureInfo.InvariantCulture);
                    return true;
                } catch (FormatException) {
                    formatted = string.Empty;
                    return false;
                }
            }

            if (!value.Number.HasValue || !IsSupportedTextNumericFormat(format)) {
                return false;
            }

            try {
                formatted = value.Number.Value.ToString(format, CultureInfo.InvariantCulture);
                return true;
            } catch (FormatException) {
                formatted = string.Empty;
                return false;
            }
        }

        private static bool LooksLikeDateTextFormat(string format) {
            bool inQuote = false;
            for (int index = 0; index < format.Length; index++) {
                char ch = format[index];
                if (ch == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                if (!inQuote && (ch == 'y' || ch == 'Y' || ch == 'd' || ch == 'D' || ch == 'h' || ch == 'H' || ch == 's' || ch == 'S')) {
                    return true;
                }
            }

            return false;
        }

        private static string ConvertExcelDateTextFormat(string format) {
            var builder = new StringBuilder(format.Length);
            bool inQuote = false;
            for (int index = 0; index < format.Length; index++) {
                char ch = format[index];
                if (ch == '"') {
                    inQuote = !inQuote;
                    builder.Append(ch);
                    continue;
                }

                if (!inQuote && (ch == 'm' || ch == 'M')) {
                    int start = index;
                    while (index + 1 < format.Length && (format[index + 1] == 'm' || format[index + 1] == 'M')) {
                        index++;
                    }

                    int count = index - start + 1;
                    bool minute = IsMinuteTextFormatToken(format, start, index);
                    builder.Append(new string(minute ? 'm' : 'M', count));
                    continue;
                }

                if (!inQuote && (ch == 'h' || ch == 'H')) {
                    int start = index;
                    while (index + 1 < format.Length && (format[index + 1] == 'h' || format[index + 1] == 'H')) {
                        index++;
                    }

                    builder.Append(new string('H', index - start + 1));
                    continue;
                }

                builder.Append(ch);
            }

            return builder.ToString();
        }

        private static bool IsMinuteTextFormatToken(string format, int start, int end) {
            char? previous = PreviousNonSpaceCharacter(format, start - 1);
            char? next = NextNonSpaceCharacter(format, end + 1);
            return previous == ':' || next == ':';
        }

        private static char? PreviousNonSpaceCharacter(string value, int start) {
            for (int index = start; index >= 0; index--) {
                if (!char.IsWhiteSpace(value[index])) {
                    return value[index];
                }
            }

            return null;
        }

        private static char? NextNonSpaceCharacter(string value, int start) {
            for (int index = start; index < value.Length; index++) {
                if (!char.IsWhiteSpace(value[index])) {
                    return value[index];
                }
            }

            return null;
        }

        private static bool IsSupportedTextNumericFormat(string format) {
            bool inQuote = false;
            foreach (char ch in format) {
                if (ch == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                if (inQuote) {
                    continue;
                }

                if (ch == '0' || ch == '#' || ch == '.' || ch == ',' || ch == '%' || ch == '$'
                    || ch == '-' || ch == '+' || ch == '(' || ch == ')' || ch == ' ') {
                    continue;
                }

                return false;
            }

            return !inQuote && format.IndexOfAny(new[] { '0', '#' }) >= 0;
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
            if (tokens.Count < 3 || tokens.Count > 6
                || !TryResolveFormulaArgument(tokens[0], out var lookupValue)
                || !TryResolveFormulaRange(tokens[1], out var lookupValues)
                || !TryResolveFormulaRange(tokens[2], out var returnValues)
                || lookupValues.Count != returnValues.Count) {
                return false;
            }

            int matchMode = 0;
            if (tokens.Count >= 5 && !TryGetWholeNumberArgument(tokens[4], out matchMode)) {
                return false;
            }

            if (matchMode != -1 && matchMode != 0 && matchMode != 1) {
                return false;
            }

            int searchMode = 1;
            if (tokens.Count >= 6
                && (!TryGetWholeNumberArgument(tokens[5], out searchMode) || (searchMode != 1 && searchMode != -1))) {
                return false;
            }

            if (TryFindLookupPosition(lookupValue, lookupValues, matchMode, searchMode, out int position)) {
                var returnValue = returnValues[position - 1];
                result = returnValue;
                return result.HasValue;
            }

            if (tokens.Count >= 4
                && TryResolveFormulaArgument(tokens[3], out FormulaArgumentValue fallback)
                && !fallback.IsUnresolvedFormula
                && fallback.HasValue) {
                result = fallback;
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

        private bool TryEvaluateStatisticalFunction(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count == 0) {
                return false;
            }

            if (function == "PERCENTILE.INC" || function == "PERCENTILE.EXC") {
                if (tokens.Count != 2
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var values)
                    || values.Count == 0
                    || !TryEvaluateFormulaOrNumeric(tokens[1], out double percentile)
                    || percentile < 0d
                    || percentile > 1d
                    || (function == "PERCENTILE.EXC" && (percentile <= 0d || percentile >= 1d))) {
                    return false;
                }

                if (function == "PERCENTILE.EXC") {
                    return TryCalculatePercentileExclusive(values, percentile, out result);
                }

                result = CalculatePercentileInclusive(values, percentile);
                return IsFinite(result);
            }

            if (function == "QUARTILE.INC" || function == "QUARTILE.EXC") {
                if (tokens.Count != 2
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var values)
                    || values.Count == 0
                    || !TryGetWholeNumberArgument(tokens[1], out int quartile)
                    || quartile < 0
                    || quartile > 4
                    || (function == "QUARTILE.EXC" && (quartile < 1 || quartile > 3))) {
                    return false;
                }

                if (function == "QUARTILE.EXC") {
                    return TryCalculatePercentileExclusive(values, quartile / 4d, out result);
                }

                result = CalculatePercentileInclusive(values, quartile / 4d);
                return IsFinite(result);
            }

            if (function == "PERCENTRANK.INC" || function == "PERCENTRANK.EXC") {
                if (tokens.Count < 2
                    || tokens.Count > 3
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var values)
                    || values.Count < 2
                    || !TryEvaluateFormulaOrNumeric(tokens[1], out double number)) {
                    return false;
                }

                int? significance = null;
                if (tokens.Count == 3) {
                    if (!TryGetWholeNumberArgument(tokens[2], out int digits) || digits < 1 || digits > 15) {
                        return false;
                    }

                    significance = digits;
                }

                return function == "PERCENTRANK.EXC"
                    ? TryEvaluatePercentRankExclusive(values, number, significance, out result)
                    : TryEvaluatePercentRankInclusive(values, number, significance, out result);
            }

            if (function == "RANK.EQ" || function == "RANK.AVG") {
                if (tokens.Count < 2
                    || tokens.Count > 3
                    || !TryEvaluateFormulaOrNumeric(tokens[0], out double number)
                    || !TryResolveFormulaArgumentNumbers(tokens[1], out var values)
                    || values.Count == 0) {
                    return false;
                }

                int order = 0;
                if (tokens.Count == 3 && !TryGetWholeNumberArgument(tokens[2], out order)) {
                    return false;
                }

                if (!values.Any(value => AreFormulaNumbersEqual(value, number))) {
                    return false;
                }

                int equalCount = values.Count(value => AreFormulaNumbersEqual(value, number));
                int betterCount = order == 0
                    ? values.Count(value => value > number && !AreFormulaNumbersEqual(value, number))
                    : values.Count(value => value < number && !AreFormulaNumbersEqual(value, number));

                double firstRank = betterCount + 1d;
                result = function == "RANK.AVG"
                    ? (firstRank + betterCount + equalCount) / 2d
                    : firstRank;
                return true;
            }

            if (function == "COVAR" || function == "COVARIANCE.P" || function == "COVARIANCE.S") {
                if (tokens.Count != 2
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var leftValues)
                    || !TryResolveFormulaArgumentNumbers(tokens[1], out var rightValues)
                    || !TryCalculateCovariance(leftValues, rightValues, sample: function == "COVARIANCE.S", out result)) {
                    return false;
                }

                return IsFinite(result);
            }

            if (function == "CORREL" || function == "SLOPE" || function == "INTERCEPT" || function == "RSQ") {
                if (tokens.Count != 2
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var knownY)
                    || !TryResolveFormulaArgumentNumbers(tokens[1], out var knownX)
                    || !TryCalculateLinearRegression(knownX, knownY, out double slope, out double intercept, out double correlation)) {
                    return false;
                }

                if (function == "CORREL") {
                    result = correlation;
                } else if (function == "SLOPE") {
                    result = slope;
                } else if (function == "INTERCEPT") {
                    result = intercept;
                } else {
                    result = correlation * correlation;
                }

                return IsFinite(result);
            }

            if (function == "FORECAST.LINEAR") {
                if (tokens.Count != 3
                    || !TryEvaluateFormulaOrNumeric(tokens[0], out double x)
                    || !TryResolveFormulaArgumentNumbers(tokens[1], out var knownY)
                    || !TryResolveFormulaArgumentNumbers(tokens[2], out var knownX)
                    || !TryCalculateLinearRegression(knownX, knownY, out double slope, out double intercept, out _)) {
                    return false;
                }

                result = intercept + slope * x;
                return IsFinite(result);
            }

            if (function == "SUMXMY2" || function == "SUMX2MY2" || function == "SUMX2PY2") {
                if (tokens.Count != 2
                    || !TryResolveFormulaArgumentNumbers(tokens[0], out var leftValues)
                    || !TryResolveFormulaArgumentNumbers(tokens[1], out var rightValues)
                    || leftValues.Count == 0
                    || leftValues.Count != rightValues.Count) {
                    return false;
                }

                result = 0;
                for (int index = 0; index < leftValues.Count; index++) {
                    double left = leftValues[index];
                    double right = rightValues[index];
                    if (function == "SUMXMY2") {
                        double delta = left - right;
                        result += delta * delta;
                    } else if (function == "SUMX2MY2") {
                        result += (left * left) - (right * right);
                    } else {
                        result += (left * left) + (right * right);
                    }
                }

                return IsFinite(result);
            }

            if (!TryResolveFormulaArguments(args, out var arguments) || arguments.Any(value => value.IsUnresolvedFormula)) {
                return false;
            }

            var numbers = arguments.Where(value => value.Number.HasValue).Select(value => value.Number!.Value).ToList();
            if (function == "AVEDEV" || function == "DEVSQ") {
                if (numbers.Count == 0) {
                    return false;
                }

                double average = numbers.Average();
                if (function == "AVEDEV") {
                    result = numbers.Sum(value => Math.Abs(value - average)) / numbers.Count;
                } else {
                    result = numbers.Sum(value => {
                        double delta = value - average;
                        return delta * delta;
                    });
                }

                return IsFinite(result);
            }

            if (function == "MODE.SNGL" || function == "MODE") {
                return TryCalculateModeSingle(numbers, out result);
            }

            if (function == "GEOMEAN") {
                if (numbers.Count == 0 || numbers.Any(value => value <= 0d)) {
                    return false;
                }

                double averageLog = numbers.Sum(Math.Log) / numbers.Count;
                result = Math.Exp(averageLog);
                return IsFinite(result);
            }

            if (function == "HARMEAN") {
                if (numbers.Count == 0 || numbers.Any(value => value <= 0d)) {
                    return false;
                }

                double reciprocalSum = numbers.Sum(value => 1d / value);
                if (Math.Abs(reciprocalSum) < double.Epsilon) {
                    return false;
                }

                result = numbers.Count / reciprocalSum;
                return IsFinite(result);
            }

            if (function == "VAR.S" || function == "STDEV.S") {
                if (numbers.Count < 2) {
                    return false;
                }

                double variance = CalculateVariance(numbers, sample: true);
                result = function == "STDEV.S" ? Math.Sqrt(variance) : variance;
                return IsFinite(result);
            }

            if (function == "VAR.P" || function == "STDEV.P") {
                if (numbers.Count < 1) {
                    return false;
                }

                double variance = CalculateVariance(numbers, sample: false);
                result = function == "STDEV.P" ? Math.Sqrt(variance) : variance;
                return IsFinite(result);
            }

            return false;
        }

        private static bool TryCalculateModeSingle(IReadOnlyList<double> numbers, out double result) {
            result = 0;
            if (numbers.Count == 0) {
                return false;
            }

            var groups = numbers
                .Select((value, index) => new { Value = value, Index = index })
                .GroupBy(item => item.Value)
                .Select(group => new { Value = group.Key, Count = group.Count(), FirstIndex = group.Min(item => item.Index) })
                .OrderByDescending(group => group.Count)
                .ThenBy(group => group.FirstIndex)
                .ToList();

            if (groups.Count == 0 || groups[0].Count < 2) {
                return false;
            }

            result = groups[0].Value;
            return IsFinite(result);
        }

        private static double CalculateVariance(IReadOnlyList<double> numbers, bool sample) {
            double average = numbers.Average();
            double sumSquares = numbers.Sum(value => {
                double delta = value - average;
                return delta * delta;
            });
            return sumSquares / (sample ? numbers.Count - 1 : numbers.Count);
        }

        private static bool TryCalculateCovariance(IReadOnlyList<double> leftValues, IReadOnlyList<double> rightValues, bool sample, out double result) {
            result = 0;
            if (leftValues.Count != rightValues.Count || leftValues.Count == 0 || (sample && leftValues.Count < 2)) {
                return false;
            }

            double leftAverage = leftValues.Average();
            double rightAverage = rightValues.Average();
            double sumProducts = 0;
            for (int index = 0; index < leftValues.Count; index++) {
                sumProducts += (leftValues[index] - leftAverage) * (rightValues[index] - rightAverage);
            }

            result = sumProducts / (sample ? leftValues.Count - 1 : leftValues.Count);
            return IsFinite(result);
        }

        private static bool TryCalculateLinearRegression(
            IReadOnlyList<double> knownX,
            IReadOnlyList<double> knownY,
            out double slope,
            out double intercept,
            out double correlation) {
            slope = 0;
            intercept = 0;
            correlation = 0;
            if (knownX.Count != knownY.Count || knownX.Count < 2) {
                return false;
            }

            double averageX = knownX.Average();
            double averageY = knownY.Average();
            double sumXy = 0;
            double sumXx = 0;
            double sumYy = 0;
            for (int index = 0; index < knownX.Count; index++) {
                double xDelta = knownX[index] - averageX;
                double yDelta = knownY[index] - averageY;
                sumXy += xDelta * yDelta;
                sumXx += xDelta * xDelta;
                sumYy += yDelta * yDelta;
            }

            if (Math.Abs(sumXx) < double.Epsilon || Math.Abs(sumYy) < double.Epsilon) {
                return false;
            }

            slope = sumXy / sumXx;
            intercept = averageY - slope * averageX;
            correlation = sumXy / Math.Sqrt(sumXx * sumYy);
            return IsFinite(slope) && IsFinite(intercept) && IsFinite(correlation);
        }

        private static bool AreFormulaNumbersEqual(double left, double right) {
            return Math.Abs(left - right) < 0.0000001d;
        }

        private static bool TryEvaluatePercentRankInclusive(IReadOnlyList<double> numbers, double number, int? significance, out double result) {
            result = 0;
            var sorted = numbers.OrderBy(value => value).ToList();
            if (number < sorted[0] || number > sorted[sorted.Count - 1]) {
                return false;
            }

            double denominator = sorted.Count - 1d;
            for (int index = 0; index < sorted.Count; index++) {
                if (AreFormulaNumbersEqual(sorted[index], number)) {
                    result = index / denominator;
                    if (significance.HasValue) {
                        result = Math.Round(result, significance.Value, MidpointRounding.AwayFromZero);
                    }

                    return IsFinite(result);
                }
            }

            for (int index = 0; index < sorted.Count - 1; index++) {
                double lower = sorted[index];
                double upper = sorted[index + 1];
                if (number <= lower || number >= upper) {
                    continue;
                }

                double fraction = (number - lower) / (upper - lower);
                result = (index + fraction) / denominator;
                if (significance.HasValue) {
                    result = Math.Round(result, significance.Value, MidpointRounding.AwayFromZero);
                }

                return IsFinite(result);
            }

            return false;
        }

        private static bool TryEvaluatePercentRankExclusive(IReadOnlyList<double> numbers, double number, int? significance, out double result) {
            result = 0;
            var sorted = numbers.OrderBy(value => value).ToList();
            if (number < sorted[0] || number > sorted[sorted.Count - 1]) {
                return false;
            }

            double denominator = sorted.Count + 1d;
            for (int index = 0; index < sorted.Count; index++) {
                if (AreFormulaNumbersEqual(sorted[index], number)) {
                    result = (index + 1d) / denominator;
                    if (significance.HasValue) {
                        result = Math.Round(result, significance.Value, MidpointRounding.AwayFromZero);
                    }

                    return IsFinite(result);
                }
            }

            for (int index = 0; index < sorted.Count - 1; index++) {
                double lower = sorted[index];
                double upper = sorted[index + 1];
                if (number <= lower || number >= upper) {
                    continue;
                }

                double fraction = (number - lower) / (upper - lower);
                result = (index + 1d + fraction) / denominator;
                if (significance.HasValue) {
                    result = Math.Round(result, significance.Value, MidpointRounding.AwayFromZero);
                }

                return IsFinite(result);
            }

            return false;
        }

        private static bool TryCalculatePercentileExclusive(IReadOnlyList<double> numbers, double percentile, out double result) {
            result = 0;
            var sorted = numbers.OrderBy(value => value).ToList();
            double rank = percentile * (sorted.Count + 1d);
            if (rank <= 1d || rank >= sorted.Count) {
                return false;
            }

            int lowerRank = (int)Math.Floor(rank);
            int upperRank = (int)Math.Ceiling(rank);
            if (lowerRank == upperRank) {
                result = sorted[lowerRank - 1];
                return IsFinite(result);
            }

            double fraction = rank - lowerRank;
            double lower = sorted[lowerRank - 1];
            double upper = sorted[upperRank - 1];
            result = lower + (upper - lower) * fraction;
            return IsFinite(result);
        }

        private static double CalculatePercentileInclusive(IReadOnlyList<double> numbers, double percentile) {
            var sorted = numbers.OrderBy(value => value).ToList();
            if (sorted.Count == 1) {
                return sorted[0];
            }

            double rank = (sorted.Count - 1) * percentile;
            int lower = (int)Math.Floor(rank);
            int upper = (int)Math.Ceiling(rank);
            if (lower == upper) {
                return sorted[lower];
            }

            double fraction = rank - lower;
            return sorted[lower] + (sorted[upper] - sorted[lower]) * fraction;
        }

        private bool TryEvaluateFinancialFunction(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);

            if (function == "NPV") {
                return TryEvaluateNpv(tokens, out result);
            }

            if (!TryResolveFormulaOrNumericArguments(tokens, out var numbers)) {
                return false;
            }

            switch (function) {
                case "PMT":
                    return TryEvaluatePmt(numbers, out result);
                case "PV":
                    return TryEvaluatePv(numbers, out result);
                case "FV":
                    return TryEvaluateFv(numbers, out result);
                case "NPER":
                    return TryEvaluateNper(numbers, out result);
                default:
                    return false;
            }
        }

        private static bool TryEvaluatePmt(IReadOnlyList<double> numbers, out double result) {
            result = 0;
            if (numbers.Count < 3 || numbers.Count > 5) {
                return false;
            }

            double rate = numbers[0];
            double periods = numbers[1];
            double presentValue = numbers[2];
            double futureValue = numbers.Count >= 4 ? numbers[3] : 0d;
            if (!TryGetPaymentType(numbers, 4, out int type) || Math.Abs(periods) < double.Epsilon) {
                return false;
            }

            if (Math.Abs(rate) < double.Epsilon) {
                result = -(presentValue + futureValue) / periods;
                return true;
            }

            double factor = Math.Pow(1d + rate, periods);
            if (!IsFinite(factor) || Math.Abs(factor - 1d) < double.Epsilon) {
                return false;
            }

            result = -(rate * (futureValue + presentValue * factor)) / ((1d + rate * type) * (factor - 1d));
            return IsFinite(result);
        }

        private static bool TryEvaluatePv(IReadOnlyList<double> numbers, out double result) {
            result = 0;
            if (numbers.Count < 3 || numbers.Count > 5) {
                return false;
            }

            double rate = numbers[0];
            double periods = numbers[1];
            double payment = numbers[2];
            double futureValue = numbers.Count >= 4 ? numbers[3] : 0d;
            if (!TryGetPaymentType(numbers, 4, out int type)) {
                return false;
            }

            if (Math.Abs(rate) < double.Epsilon) {
                result = -(futureValue + payment * periods);
                return true;
            }

            double factor = Math.Pow(1d + rate, periods);
            if (!IsFinite(factor) || Math.Abs(factor) < double.Epsilon) {
                return false;
            }

            result = -(futureValue + payment * (1d + rate * type) * ((factor - 1d) / rate)) / factor;
            return IsFinite(result);
        }

        private static bool TryEvaluateFv(IReadOnlyList<double> numbers, out double result) {
            result = 0;
            if (numbers.Count < 3 || numbers.Count > 5) {
                return false;
            }

            double rate = numbers[0];
            double periods = numbers[1];
            double payment = numbers[2];
            double presentValue = numbers.Count >= 4 ? numbers[3] : 0d;
            if (!TryGetPaymentType(numbers, 4, out int type)) {
                return false;
            }

            if (Math.Abs(rate) < double.Epsilon) {
                result = -(presentValue + payment * periods);
                return true;
            }

            double factor = Math.Pow(1d + rate, periods);
            if (!IsFinite(factor)) {
                return false;
            }

            result = -(presentValue * factor + payment * (1d + rate * type) * ((factor - 1d) / rate));
            return IsFinite(result);
        }

        private static bool TryEvaluateNper(IReadOnlyList<double> numbers, out double result) {
            result = 0;
            if (numbers.Count < 3 || numbers.Count > 5) {
                return false;
            }

            double rate = numbers[0];
            double payment = numbers[1];
            double presentValue = numbers[2];
            double futureValue = numbers.Count >= 4 ? numbers[3] : 0d;
            if (!TryGetPaymentType(numbers, 4, out int type)) {
                return false;
            }

            if (Math.Abs(rate) < double.Epsilon) {
                if (Math.Abs(payment) < double.Epsilon) {
                    return false;
                }

                result = -(presentValue + futureValue) / payment;
                return IsFinite(result);
            }

            double adjustedPayment = payment * (1d + rate * type);
            double numerator = adjustedPayment - futureValue * rate;
            double denominator = presentValue * rate + adjustedPayment;
            if (Math.Abs(denominator) < double.Epsilon || 1d + rate <= 0d) {
                return false;
            }

            double ratio = numerator / denominator;
            if (ratio <= 0d) {
                return false;
            }

            result = Math.Log(ratio) / Math.Log(1d + rate);
            return IsFinite(result);
        }

        private bool TryEvaluateNpv(IReadOnlyList<string> tokens, out double result) {
            result = 0;
            if (tokens.Count < 2 || !TryEvaluateFormulaOrNumeric(tokens[0], out double rate) || 1d + rate == 0d) {
                return false;
            }

            int period = 1;
            for (int index = 1; index < tokens.Count; index++) {
                if (!TryResolveFormulaArgumentNumbers(tokens[index], out var values) || values.Count == 0) {
                    return false;
                }

                foreach (double value in values) {
                    double divisor = Math.Pow(1d + rate, period);
                    if (!IsFinite(divisor) || Math.Abs(divisor) < double.Epsilon) {
                        return false;
                    }

                    result += value / divisor;
                    period++;
                }
            }

            return IsFinite(result);
        }

        private static bool TryGetPaymentType(IReadOnlyList<double> numbers, int index, out int type) {
            type = 0;
            if (numbers.Count <= index) {
                return true;
            }

            if (!TryGetWholeNumber(numbers[index], out type)) {
                return false;
            }

            return type == 0 || type == 1;
        }

        private static bool IsFinite(double value) {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static bool TryEvaluateMRound(double number, double multiple, out double result) {
            result = 0;
            if (!IsFinite(number) || !IsFinite(multiple)) {
                return false;
            }

            if (Math.Abs(multiple) < double.Epsilon) {
                result = 0;
                return true;
            }

            if (Math.Sign(number) != 0 && Math.Sign(multiple) != 0 && Math.Sign(number) != Math.Sign(multiple)) {
                return false;
            }

            result = Math.Round(number / multiple, 0, MidpointRounding.AwayFromZero) * multiple;
            return IsFinite(result);
        }

        private static bool TryEvaluateMathRoundFunction(string function, IReadOnlyList<double> numbers, out double result) {
            result = 0;
            double number = numbers[0];
            double significance = numbers.Count >= 2 ? Math.Abs(numbers[1]) : 1d;
            double mode = numbers.Count >= 3 ? numbers[2] : 0d;
            if (!IsFinite(number) || !IsFinite(significance) || !IsFinite(mode)) {
                return false;
            }

            if (Math.Abs(significance) < double.Epsilon) {
                result = 0;
                return true;
            }

            double quotient = number / significance;
            if (number >= 0) {
                result = (function == "CEILING.MATH" ? Math.Ceiling(quotient) : Math.Floor(quotient)) * significance;
                return IsFinite(result);
            }

            bool useAwayFromZeroForNegative = function == "CEILING.MATH"
                ? Math.Abs(mode) >= double.Epsilon
                : Math.Abs(mode) < double.Epsilon;
            result = (useAwayFromZeroForNegative ? Math.Floor(quotient) : Math.Ceiling(quotient)) * significance;
            return IsFinite(result);
        }

        private bool TryEvaluateErrorFallback(string function, string args, out double result) {
            result = 0;
            bool success = function == "IFNA"
                ? TryEvaluateIfNaValue(args, out FormulaArgumentValue value)
                : TryEvaluateIfErrorValue(args, out value);
            if (!success || !value.Number.HasValue) {
                return false;
            }

            result = value.Number.Value;
            return true;
        }

        private bool TryEvaluateIf(string args, out double result) {
            result = 0;
            if (!TryEvaluateIfValue(args, out FormulaArgumentValue value) || !value.Number.HasValue) {
                return false;
            }

            result = value.Number.Value;
            return true;
        }

        private bool TryEvaluateIfErrorValue(string args, out FormulaArgumentValue result) {
            return TryEvaluateErrorFallbackValue(args, _ => true, out result);
        }

        private bool TryEvaluateIfNaValue(string args, out FormulaArgumentValue result) {
            return TryEvaluateErrorFallbackValue(args, errorCode => string.Equals(errorCode, "#N/A", StringComparison.OrdinalIgnoreCase), out result);
        }

        private bool TryEvaluateErrorFallbackValue(string args, Func<string, bool> shouldUseFallback, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 2) {
                return false;
            }

            if (!TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue candidate)
                || candidate.IsUnresolvedFormula
                || !candidate.HasValue) {
                if (!TryInferFormulaErrorArgument(tokens[0], out string inferredErrorCode)) {
                    return false;
                }

                candidate = FormulaArgumentValue.Error(inferredErrorCode);
            }

            if (!candidate.IsError || !shouldUseFallback(candidate.ErrorCode ?? "#VALUE!")) {
                result = candidate;
                return true;
            }

            return TryResolveFormulaArgument(tokens[1], out result)
                && !result.IsUnresolvedFormula
                && result.HasValue;
        }

        private bool TryEvaluateIfValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 3
                || !TryEvaluateCondition(tokens[0], out bool condition)
                || !TryResolveFormulaArgument(tokens[condition ? 1 : 2], out result)
                || result.IsUnresolvedFormula
                || !result.HasValue) {
                return false;
            }

            return true;
        }

        private bool TryEvaluateIfsValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 2 || tokens.Count % 2 != 0) {
                return false;
            }

            for (int index = 0; index < tokens.Count; index += 2) {
                if (!TryEvaluateCondition(tokens[index], out bool condition)) {
                    return false;
                }

                if (!condition) {
                    continue;
                }

                return TryResolveFormulaArgument(tokens[index + 1], out result)
                    && !result.IsUnresolvedFormula
                    && result.HasValue;
            }

            return false;
        }

        private bool TryEvaluateSwitchValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 3) {
                return false;
            }

            if (!TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue expression)
                || expression.IsUnresolvedFormula
                || !expression.HasValue) {
                return false;
            }

            int pairEnd = tokens.Count;
            bool hasDefault = tokens.Count % 2 == 0;
            if (hasDefault) {
                pairEnd--;
            }

            for (int index = 1; index < pairEnd; index += 2) {
                if (!TryResolveFormulaArgument(tokens[index], out FormulaArgumentValue candidate)
                    || candidate.IsUnresolvedFormula
                    || !candidate.HasValue) {
                    return false;
                }

                if (!FormulaValuesEqual(expression, candidate)) {
                    continue;
                }

                return TryResolveFormulaArgument(tokens[index + 1], out result)
                    && !result.IsUnresolvedFormula
                    && result.HasValue;
            }

            return hasDefault
                && TryResolveFormulaArgument(tokens[tokens.Count - 1], out result)
                && !result.IsUnresolvedFormula
                && result.HasValue;
        }

        private bool TryEvaluateChooseValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 2 || !TryGetWholeNumberArgument(tokens[0], out int index)) {
                return false;
            }

            if (index < 1 || index >= tokens.Count) {
                return false;
            }

            return TryResolveFormulaArgument(tokens[index], out result)
                && !result.IsUnresolvedFormula
                && result.HasValue;
        }

        private bool TryEvaluateInfoFunction(string function, string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (function == "ISFORMULA") {
                if (tokens.Count != 1 || !TryResolveFormulaReferenceArgument(tokens[0], out ExcelSheet referenceSheet, out int referenceRow, out int referenceColumn)) {
                    return false;
                }

                bool isFormula = referenceSheet.TryGetExistingCell(referenceRow, referenceColumn)?.CellFormula != null;
                result = new FormulaArgumentValue(isFormula ? 1d : 0d, isFormula ? "1" : "0");
                return true;
            }

            if (tokens.Count != 1 || !TryResolveInfoArgument(tokens[0], out FormulaArgumentValue value)) {
                return false;
            }

            bool matches;
            switch (function) {
                case "ISBLANK":
                    matches = !value.HasValue;
                    break;
                case "ISNUMBER":
                    matches = value.Number.HasValue;
                    break;
                case "ISTEXT":
                    matches = value.Text != null && !value.Number.HasValue;
                    break;
                case "ISERROR":
                    matches = value.IsError;
                    break;
                case "ISERR":
                    matches = value.IsError && !string.Equals(value.ErrorCode, "#N/A", StringComparison.OrdinalIgnoreCase);
                    break;
                case "ISNA":
                    matches = value.IsError && string.Equals(value.ErrorCode, "#N/A", StringComparison.OrdinalIgnoreCase);
                    break;
                default:
                    return false;
            }

            result = new FormulaArgumentValue(matches ? 1d : 0d, matches ? "1" : "0");
            return true;
        }

        private bool TryEvaluateReferenceShapeFunction(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 1) {
                return false;
            }

            if (function == "ROW" || function == "COLUMN") {
                if (!TryResolveFormulaReferenceArgument(tokens[0], out _, out int row, out int column)) {
                    return false;
                }

                result = function == "ROW" ? row : column;
                return true;
            }

            if (!TryResolveFormulaRangeReference(tokens[0], out _, out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }

            result = function == "ROWS"
                ? r2 - r1 + 1
                : c2 - c1 + 1;
            return result > 0;
        }

        private bool TryGetFormulaTextArgument(string token, out string formulaText) {
            formulaText = string.Empty;
            if (!TryResolveFormulaReferenceArgument(token, out ExcelSheet referenceSheet, out int row, out int column)) {
                return false;
            }

            formulaText = referenceSheet.TryGetExistingCell(row, column)?.CellFormula?.Text ?? string.Empty;
            return formulaText.Length > 0;
        }

        private bool TryResolveInfoArgument(string token, out FormulaArgumentValue value) {
            if (TryResolveFormulaArgument(token, out value)) {
                return true;
            }

            if (TryInferFormulaErrorArgument(token, out string errorCode)) {
                value = FormulaArgumentValue.Error(errorCode);
                return true;
            }

            value = default;
            return false;
        }

        private bool TryInferFormulaErrorArgument(string token, out string errorCode) {
            string trimmed = token.Trim();
            if (TryParseFormulaErrorLiteral(trimmed, out errorCode)) {
                return true;
            }

            Match functionMatch = SimpleFunctionFormulaRegex.Match(trimmed);
            if (functionMatch.Success) {
                string function = functionMatch.Groups[1].Value.ToUpperInvariant();
                errorCode = function == "MATCH" || function == "XMATCH" || function == "XLOOKUP" || function == "VLOOKUP" || function == "HLOOKUP"
                    ? "#N/A"
                    : "#VALUE!";
                return true;
            }

            var binaryMatch = SimpleBinaryFormulaRegex.Match(trimmed);
            if (binaryMatch.Success
                && string.Equals(binaryMatch.Groups[2].Value, "/", StringComparison.Ordinal)
                && TryResolveNumericOperand(binaryMatch.Groups[3].Value, out double divisor)
                && Math.Abs(divisor) < double.Epsilon) {
                errorCode = "#DIV/0!";
                return true;
            }

            errorCode = string.Empty;
            return false;
        }

        private bool TryEvaluateFormulaOrNumeric(string token, out double result) {
            if (TryEvaluateFormula(token, out result) || TryResolveNumericOperand(token, out result)) {
                return true;
            }

            if (TryResolveFormulaArgument(token, out FormulaArgumentValue value) && value.Number.HasValue) {
                result = value.Number.Value;
                return true;
            }

            return false;
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

        private bool TryEvaluateCountBlank(string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count != 1 || !TryResolveFormulaRange(tokens[0], out var values)) {
                return false;
            }

            result = values.Count(IsFormulaBlankValue);
            return true;
        }

        private bool TryEvaluateSubtotal(string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 2 || !TryGetWholeNumberArgument(tokens[0], out int functionCode)) {
                return false;
            }

            functionCode %= 100;
            if (functionCode != 1 && functionCode != 2 && functionCode != 3 && functionCode != 4 && functionCode != 5 && functionCode != 9) {
                return false;
            }

            var values = new List<FormulaArgumentValue>();
            for (int index = 1; index < tokens.Count; index++) {
                if (!TryResolveFormulaRange(tokens[index], out var rangeValues)) {
                    return false;
                }

                values.AddRange(rangeValues);
            }

            if (values.Count == 0) {
                return false;
            }

            if (functionCode == 3) {
                result = values.Count(value => !value.IsUnresolvedFormula && !IsFormulaBlankValue(value));
                return true;
            }

            var numbers = values.Where(value => value.Number.HasValue).Select(value => value.Number!.Value).ToList();
            if (functionCode == 2) {
                result = numbers.Count;
                return true;
            }

            if (functionCode == 9) {
                result = numbers.Sum();
                return true;
            }

            if (numbers.Count == 0) {
                return false;
            }

            if (functionCode == 1) {
                result = numbers.Average();
            } else if (functionCode == 4) {
                result = numbers.Max();
            } else {
                result = numbers.Min();
            }

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

            if (function == "WORKDAY" || function == "WORKDAY.INTL") {
                return TryEvaluateWorkday(function, tokens, out result);
            }

            if (function == "DATEVALUE" || function == "TIMEVALUE") {
                return TryEvaluateDateTimeTextValue(function, tokens, out result);
            }

            if (function == "DATEDIF") {
                return TryEvaluateDateDif(tokens, out result);
            }

            if (function == "YEARFRAC") {
                return TryEvaluateYearFrac(tokens, out result);
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

            if (function == "DAYS360") {
                if (tokens.Count < 2
                    || tokens.Count > 3
                    || !TryEvaluateFormulaOrNumeric(tokens[0], out double startSerial)
                    || !TryEvaluateFormulaOrNumeric(tokens[1], out double endSerial)
                    || !TryGetDateFromSerial(startSerial, out DateTime startDate)
                    || !TryGetDateFromSerial(endSerial, out DateTime endDate)) {
                    return false;
                }

                bool europeanMethod = false;
                if (tokens.Count == 3 && !TryResolveBooleanArgument(tokens[2], out europeanMethod)) {
                    return false;
                }

                result = europeanMethod ? Days360European(startDate, endDate) : Days360Us(startDate, endDate);
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

            if (function == "WEEKNUM" || function == "ISOWEEKNUM") {
                if (numbers.Count < 1
                    || numbers.Count > 2
                    || (function == "ISOWEEKNUM" && numbers.Count != 1)
                    || !TryGetDateFromSerial(numbers[0], out DateTime date)) {
                    return false;
                }

                if (function == "ISOWEEKNUM") {
                    result = GetIsoWeekNumber(date);
                    return true;
                }

                int returnType = 1;
                if (numbers.Count == 2 && !TryGetWholeNumber(numbers[1], out returnType)) {
                    return false;
                }

                return TryGetWeekStartDay(returnType, out DayOfWeek weekStart)
                    && TryGetWeekNumber(date, weekStart, returnType == 21, out result);
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

        private bool TryEvaluateYearFrac(IReadOnlyList<string> tokens, out double result) {
            result = 0;
            if (tokens.Count < 2
                || tokens.Count > 3
                || !TryEvaluateFormulaOrNumeric(tokens[0], out double startSerial)
                || !TryEvaluateFormulaOrNumeric(tokens[1], out double endSerial)
                || !TryGetDateFromSerial(startSerial, out DateTime startDate)
                || !TryGetDateFromSerial(endSerial, out DateTime endDate)
                || endDate < startDate) {
                return false;
            }

            int basis = 0;
            if (tokens.Count == 3 && !TryGetWholeNumberArgument(tokens[2], out basis)) {
                return false;
            }

            switch (basis) {
                case 0:
                    result = Days360Us(startDate, endDate) / 360d;
                    return true;
                case 1:
                    result = ActualActualYearFraction(startDate, endDate);
                    return true;
                case 2:
                    result = (endDate - startDate).TotalDays / 360d;
                    return true;
                case 3:
                    result = (endDate - startDate).TotalDays / 365d;
                    return true;
                case 4:
                    result = Days360European(startDate, endDate) / 360d;
                    return true;
                default:
                    return false;
            }
        }

        private static int Days360Us(DateTime startDate, DateTime endDate) {
            int startDay = startDate.Day;
            int endDay = endDate.Day;

            if (startDay == 31 || IsLastDayOfFebruary(startDate)) {
                startDay = 30;
            }

            if (endDay == 31 && startDay >= 30) {
                endDay = 30;
            }

            return ((endDate.Year - startDate.Year) * 360)
                + ((endDate.Month - startDate.Month) * 30)
                + endDay - startDay;
        }

        private static int Days360European(DateTime startDate, DateTime endDate) {
            int startDay = Math.Min(startDate.Day, 30);
            int endDay = Math.Min(endDate.Day, 30);
            return ((endDate.Year - startDate.Year) * 360)
                + ((endDate.Month - startDate.Month) * 30)
                + endDay - startDay;
        }

        private static bool TryGetWeekStartDay(int returnType, out DayOfWeek weekStart) {
            switch (returnType) {
                case 1:
                case 17:
                    weekStart = DayOfWeek.Sunday;
                    return true;
                case 2:
                case 11:
                case 21:
                    weekStart = DayOfWeek.Monday;
                    return true;
                case 12:
                    weekStart = DayOfWeek.Tuesday;
                    return true;
                case 13:
                    weekStart = DayOfWeek.Wednesday;
                    return true;
                case 14:
                    weekStart = DayOfWeek.Thursday;
                    return true;
                case 15:
                    weekStart = DayOfWeek.Friday;
                    return true;
                case 16:
                    weekStart = DayOfWeek.Saturday;
                    return true;
                default:
                    weekStart = DayOfWeek.Sunday;
                    return false;
            }
        }

        private static bool TryGetWeekNumber(DateTime date, DayOfWeek weekStart, bool isoSystem, out double result) {
            if (isoSystem) {
                result = GetIsoWeekNumber(date);
                return true;
            }

            DateTime firstDay = new DateTime(date.Year, 1, 1);
            DateTime firstWeekStart = firstDay.AddDays(-GetDayOffset(firstDay.DayOfWeek, weekStart));
            result = Math.Floor((date.Date - firstWeekStart).TotalDays / 7d) + 1d;
            return result >= 1d && result <= 54d;
        }

        private static int GetIsoWeekNumber(DateTime date) {
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(date);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday) {
                date = date.AddDays(3);
            }

            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(
                date,
                CalendarWeekRule.FirstFourDayWeek,
                DayOfWeek.Monday);
        }

        private static int GetDayOffset(DayOfWeek day, DayOfWeek weekStart) {
            int offset = (int)day - (int)weekStart;
            return offset < 0 ? offset + 7 : offset;
        }

        private static double ActualActualYearFraction(DateTime startDate, DateTime endDate) {
            if (startDate == endDate) {
                return 0d;
            }

            if (startDate.Year == endDate.Year) {
                return (endDate - startDate).TotalDays / DaysInYear(startDate.Year);
            }

            double fraction = (new DateTime(startDate.Year + 1, 1, 1) - startDate).TotalDays / DaysInYear(startDate.Year);
            for (int year = startDate.Year + 1; year < endDate.Year; year++) {
                fraction += 1d;
            }

            fraction += (endDate - new DateTime(endDate.Year, 1, 1)).TotalDays / DaysInYear(endDate.Year);
            return fraction;
        }

        private static int DaysInYear(int year) {
            return DateTime.IsLeapYear(year) ? 366 : 365;
        }

        private static bool IsLastDayOfFebruary(DateTime date) {
            return date.Month == 2 && date.Day == DateTime.DaysInMonth(date.Year, 2);
        }

        private bool TryEvaluateDateDif(IReadOnlyList<string> tokens, out double result) {
            result = 0;
            if (tokens.Count != 3
                || !TryEvaluateFormulaOrNumeric(tokens[0], out double startSerial)
                || !TryEvaluateFormulaOrNumeric(tokens[1], out double endSerial)
                || !TryGetDateFromSerial(startSerial, out DateTime startDate)
                || !TryGetDateFromSerial(endSerial, out DateTime endDate)
                || endDate < startDate
                || !TryResolveTextArgument(tokens[2], out string unit)) {
                return false;
            }

            switch (unit.ToUpperInvariant()) {
                case "D":
                    result = (endDate - startDate).TotalDays;
                    return true;
                case "M":
                    result = GetCompletedMonths(startDate, endDate);
                    return true;
                case "Y":
                    result = GetCompletedYears(startDate, endDate);
                    return true;
                case "YM":
                    result = GetRemainingCompletedMonthsAfterYears(startDate, endDate);
                    return true;
                case "YD":
                    result = GetDaysAfterLastAnniversary(startDate, endDate);
                    return true;
                case "MD":
                    result = GetRemainingDaysAfterMonths(startDate, endDate);
                    return true;
                default:
                    return false;
            }
        }

        private static int GetCompletedYears(DateTime startDate, DateTime endDate) {
            int years = endDate.Year - startDate.Year;
            if (endDate < AddYearsClamped(startDate, years)) {
                years--;
            }

            return years;
        }

        private static int GetCompletedMonths(DateTime startDate, DateTime endDate) {
            int months = (endDate.Year - startDate.Year) * 12 + endDate.Month - startDate.Month;
            if (endDate.Day < startDate.Day) {
                months--;
            }

            return months;
        }

        private static int GetRemainingCompletedMonthsAfterYears(DateTime startDate, DateTime endDate) {
            int years = GetCompletedYears(startDate, endDate);
            DateTime anniversary = AddYearsClamped(startDate, years);
            int months = endDate.Month - anniversary.Month;
            if (months < 0) {
                months += 12;
            }

            if (endDate.Day < anniversary.Day) {
                months--;
                if (months < 0) {
                    months += 12;
                }
            }

            return months;
        }

        private static int GetDaysAfterLastAnniversary(DateTime startDate, DateTime endDate) {
            DateTime anniversary = CreateClampedDate(endDate.Year, startDate.Month, startDate.Day);
            if (anniversary > endDate) {
                anniversary = CreateClampedDate(endDate.Year - 1, startDate.Month, startDate.Day);
            }

            return (int)(endDate - anniversary).TotalDays;
        }

        private static int GetRemainingDaysAfterMonths(DateTime startDate, DateTime endDate) {
            if (endDate.Day >= startDate.Day) {
                return endDate.Day - startDate.Day;
            }

            DateTime previousMonth = endDate.AddMonths(-1);
            int daysInPreviousMonth = DateTime.DaysInMonth(previousMonth.Year, previousMonth.Month);
            return endDate.Day + daysInPreviousMonth - startDate.Day;
        }

        private static DateTime AddYearsClamped(DateTime date, int years) {
            return CreateClampedDate(date.Year + years, date.Month, date.Day);
        }

        private static DateTime CreateClampedDate(int year, int month, int day) {
            int clampedDay = Math.Min(day, DateTime.DaysInMonth(year, month));
            return new DateTime(year, month, clampedDay);
        }

        private bool TryEvaluateWorkday(string function, IReadOnlyList<string> tokens, out double result) {
            result = 0;
            int maxTokens = function == "WORKDAY.INTL" ? 4 : 3;
            if (tokens.Count < 2 || tokens.Count > maxTokens
                || !TryEvaluateFormulaOrNumeric(tokens[0], out double startSerial)
                || !TryGetWholeNumberArgument(tokens[1], out int days)
                || !TryGetDateFromSerial(startSerial, out DateTime current)) {
                return false;
            }

            bool[] weekendMask = DefaultWeekendMask();
            int holidayIndex = 2;
            if (function == "WORKDAY.INTL") {
                holidayIndex = 3;
                if (tokens.Count >= 3 && !TryResolveWeekendMask(tokens[2], weekendMask)) {
                    return false;
                }
            }

            var holidays = new HashSet<DateTime>();
            if (tokens.Count > holidayIndex && !TryResolveHolidayDates(tokens[holidayIndex], holidays)) {
                return false;
            }

            if (days == 0) {
                result = current.ToOADate();
                return true;
            }

            int direction = days > 0 ? 1 : -1;
            int remaining = Math.Abs(days);
            while (remaining > 0) {
                current = current.AddDays(direction);
                if (!IsMaskedWeekend(current.DayOfWeek, weekendMask) && !holidays.Contains(current.Date)) {
                    remaining--;
                }
            }

            result = current.ToOADate();
            return true;
        }

        private bool TryEvaluateDateTimeTextValue(string function, IReadOnlyList<string> tokens, out double result) {
            result = 0;
            if (tokens.Count != 1 || !TryResolveTextArgument(tokens[0], out string text)) {
                return false;
            }

            text = text.Trim();
            if (string.IsNullOrEmpty(text)) {
                return false;
            }

            if (function == "DATEVALUE") {
                if (!TryParseFormulaDateText(text, out DateTime date)) {
                    return false;
                }

                result = date.Date.ToOADate();
                return true;
            }

            if (!TryParseFormulaTimeText(text, out TimeSpan time)) {
                return false;
            }

            result = time.TotalDays;
            return true;
        }

        private static bool TryParseFormulaDateText(string text, out DateTime date) {
            string[] exactFormats = {
                "yyyy-MM-dd",
                "yyyy-M-d",
                "yyyy/MM/dd",
                "yyyy/M/d",
                "MM/dd/yyyy",
                "M/d/yyyy",
                "dd-MMM-yyyy",
                "d-MMM-yyyy",
                "MMM d yyyy",
                "MMMM d yyyy"
            };

            return DateTime.TryParseExact(text, exactFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out date)
                || DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out date);
        }

        private static bool TryParseFormulaTimeText(string text, out TimeSpan time) {
            string[] exactFormats = {
                "H:mm",
                "HH:mm",
                "H:mm:ss",
                "HH:mm:ss",
                "h:mm tt",
                "hh:mm tt",
                "h:mm:ss tt",
                "hh:mm:ss tt"
            };

            if (DateTime.TryParseExact(text, exactFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime exactTime)) {
                time = exactTime.TimeOfDay;
                return true;
            }

            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime parsedTime)) {
                time = parsedTime.TimeOfDay;
                return true;
            }

            time = default;
            return false;
        }

        private bool TryResolveWeekendMask(string token, bool[] weekendMask) {
            string trimmed = token.Trim();
            if (TryResolveTextArgument(trimmed, out string maskText) && TryParseWeekendTextMask(maskText, weekendMask)) {
                return true;
            }

            if (!TryGetWholeNumberArgument(trimmed, out int weekendCode)) {
                return false;
            }

            return TryApplyWeekendCode(weekendCode, weekendMask);
        }

        private static bool[] DefaultWeekendMask() {
            var weekendMask = new bool[7];
            weekendMask[(int)DayOfWeek.Saturday] = true;
            weekendMask[(int)DayOfWeek.Sunday] = true;
            return weekendMask;
        }

        private static bool TryParseWeekendTextMask(string text, bool[] weekendMask) {
            if (text.Length != 7 || text.Any(ch => ch != '0' && ch != '1') || text.All(ch => ch == '1')) {
                return false;
            }

            Array.Clear(weekendMask, 0, weekendMask.Length);
            for (int index = 0; index < text.Length; index++) {
                DayOfWeek day = index == 6 ? DayOfWeek.Sunday : (DayOfWeek)(index + 1);
                weekendMask[(int)day] = text[index] == '1';
            }

            return true;
        }

        private static bool TryApplyWeekendCode(int weekendCode, bool[] weekendMask) {
            Array.Clear(weekendMask, 0, weekendMask.Length);
            switch (weekendCode) {
                case 1:
                    weekendMask[(int)DayOfWeek.Saturday] = true;
                    weekendMask[(int)DayOfWeek.Sunday] = true;
                    return true;
                case 2:
                    weekendMask[(int)DayOfWeek.Sunday] = true;
                    weekendMask[(int)DayOfWeek.Monday] = true;
                    return true;
                case 3:
                    weekendMask[(int)DayOfWeek.Monday] = true;
                    weekendMask[(int)DayOfWeek.Tuesday] = true;
                    return true;
                case 4:
                    weekendMask[(int)DayOfWeek.Tuesday] = true;
                    weekendMask[(int)DayOfWeek.Wednesday] = true;
                    return true;
                case 5:
                    weekendMask[(int)DayOfWeek.Wednesday] = true;
                    weekendMask[(int)DayOfWeek.Thursday] = true;
                    return true;
                case 6:
                    weekendMask[(int)DayOfWeek.Thursday] = true;
                    weekendMask[(int)DayOfWeek.Friday] = true;
                    return true;
                case 7:
                    weekendMask[(int)DayOfWeek.Friday] = true;
                    weekendMask[(int)DayOfWeek.Saturday] = true;
                    return true;
                default:
                    if (weekendCode >= 11 && weekendCode <= 17) {
                        DayOfWeek singleWeekendDay = weekendCode == 11 ? DayOfWeek.Sunday : (DayOfWeek)(weekendCode - 11);
                        weekendMask[(int)singleWeekendDay] = true;
                        return true;
                    }

                    return false;
            }
        }

        private static bool IsMaskedWeekend(DayOfWeek day, bool[] weekendMask) {
            return weekendMask[(int)day];
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

            if (function == "MINIFS") {
                result = numbers.Min();
            } else if (function == "MAXIFS") {
                result = numbers.Max();
            } else {
                result = numbers.Average();
            }
            return true;
        }

        private bool TryEvaluateCondition(string condition, out bool result) {
            result = false;
            if (TryResolveBooleanArgument(condition, out bool booleanResult)) {
                result = booleanResult;
                return true;
            }

            if (TrySplitFormulaComparison(condition, out string leftToken, out string comparisonOperator, out string rightToken)
                && TryResolveFormulaArgument(leftToken, out FormulaArgumentValue leftValue)
                && TryResolveFormulaArgument(rightToken, out FormulaArgumentValue rightValue)
                && !leftValue.IsUnresolvedFormula
                && !rightValue.IsUnresolvedFormula
                && leftValue.HasValue
                && rightValue.HasValue) {
                return TryCompareFormulaValues(leftValue, comparisonOperator, rightValue, out result);
            }

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

        private static bool TrySplitFormulaComparison(string condition, out string left, out string comparisonOperator, out string right) {
            left = string.Empty;
            comparisonOperator = string.Empty;
            right = string.Empty;
            string value = condition.Trim();
            int depth = 0;
            int bracketDepth = 0;
            bool inString = false;

            for (int index = 0; index < value.Length; index++) {
                char ch = value[index];
                if (ch == '"') {
                    if (inString && index + 1 < value.Length && value[index + 1] == '"') {
                        index++;
                        continue;
                    }

                    inString = !inString;
                    continue;
                }

                if (inString) {
                    continue;
                }

                if (ch == '(') {
                    depth++;
                    continue;
                }

                if (ch == ')') {
                    depth--;
                    if (depth < 0) {
                        return false;
                    }

                    continue;
                }

                if (ch == '[') {
                    bracketDepth++;
                    continue;
                }

                if (ch == ']') {
                    bracketDepth--;
                    if (bracketDepth < 0) {
                        return false;
                    }

                    continue;
                }

                if (depth != 0 || bracketDepth != 0) {
                    continue;
                }

                string? op = null;
                if (index + 1 < value.Length) {
                    string pair = value.Substring(index, 2);
                    if (pair == ">=" || pair == "<=" || pair == "<>") {
                        op = pair;
                    }
                }

                if (op == null && (ch == '=' || ch == '>' || ch == '<')) {
                    op = ch.ToString();
                }

                if (op == null) {
                    continue;
                }

                left = value.Substring(0, index).Trim();
                comparisonOperator = op;
                right = value.Substring(index + op.Length).Trim();
                return left.Length > 0 && right.Length > 0;
            }

            return false;
        }

        private static bool TryCompareFormulaValues(FormulaArgumentValue left, string comparisonOperator, FormulaArgumentValue right, out bool result) {
            result = false;
            int comparison;
            if (left.Number.HasValue && right.Number.HasValue) {
                double delta = left.Number.Value - right.Number.Value;
                comparison = Math.Abs(delta) < 0.0000001 ? 0 : delta < 0 ? -1 : 1;
            } else {
                string leftText = left.Text ?? left.Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
                string rightText = right.Text ?? right.Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
                comparison = string.Compare(leftText, rightText, StringComparison.OrdinalIgnoreCase);
            }

            switch (comparisonOperator) {
                case ">":
                    result = comparison > 0;
                    return true;
                case "<":
                    result = comparison < 0;
                    return true;
                case ">=":
                    result = comparison >= 0;
                    return true;
                case "<=":
                    result = comparison <= 0;
                    return true;
                case "=":
                    result = comparison == 0;
                    return true;
                case "<>":
                    result = comparison != 0;
                    return true;
                default:
                    return false;
            }
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

        private bool TryParseCriteria(string token, out FormulaCriteria criteria) {
            string value = token.Trim();

            if (TryResolveFormulaArgument(value, out FormulaArgumentValue criteriaValue) && !criteriaValue.IsUnresolvedFormula && criteriaValue.HasValue) {
                value = FormulaValueToText(criteriaValue);
            } else if (value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"') {
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

        private static bool IsFormulaBlankValue(FormulaArgumentValue value) {
            return !value.IsUnresolvedFormula && !value.Number.HasValue && string.IsNullOrEmpty(value.Text);
        }

        private static bool TryGetSupportedDecimalPlaces(double value, out int digits) {
            digits = (int)Math.Round(value, MidpointRounding.AwayFromZero);
            return Math.Abs(value - digits) < 0.0000001 && digits >= -15 && digits <= 15;
        }

        private static double RoundAtDigits(double value, int digits, MidpointRounding mode) {
            if (digits >= 0) {
                return Math.Round(value, digits, mode);
            }

            double factor = Math.Pow(10, -digits);
            return Math.Round(value / factor, 0, mode) * factor;
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

        private static bool TryGetDateTimeFromSerial(double value, out DateTime date) {
            date = default;
            try {
                date = DateTime.FromOADate(value);
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

        private IReadOnlyList<string> GetFormulaDependencies(string formula) {
            if (string.IsNullOrWhiteSpace(formula)) {
                return Array.Empty<string>();
            }

            try {
                string searchableFormula = MaskFormulaStringLiterals(formula);
                return FormulaReferenceRegex.Matches(searchableFormula)
                    .Cast<Match>()
                    .Select(match => match.Groups["reference"].Value)
                    .Where(reference => !string.IsNullOrWhiteSpace(reference))
                    .Select(NormalizeFormulaDependencyReference)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(reference => reference, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            } catch (RegexMatchTimeoutException) {
                return Array.Empty<string>();
            }
        }

        private IReadOnlyList<string> GetFormulaDependencyIssues(string formula, string? sourceCellReference, IReadOnlyList<string> dependencies) {
            if (dependencies.Count == 0) {
                return Array.Empty<string>();
            }

            var issues = new List<string>();
            string? sourceReference = NormalizeFormulaCellReference(sourceCellReference);
            foreach (string dependency in dependencies) {
                if (!TryResolveFormulaRangeReference(dependency, out ExcelSheet dependencySheet, out int r1, out int c1, out int r2, out int c2)) {
                    issues.Add($"Cannot resolve dependency '{dependency}'.");
                    continue;
                }

                if (sourceReference != null
                    && string.Equals(dependencySheet.Name, Name, StringComparison.OrdinalIgnoreCase)
                    && TryParseCellReference(sourceReference, out int sourceRow, out int sourceColumn)
                    && sourceRow >= r1 && sourceRow <= r2 && sourceColumn >= c1 && sourceColumn <= c2) {
                    issues.Add($"Dependency '{dependency}' references its own formula cell.");
                }

                foreach (Cell dependencyCell in dependencySheet.WorksheetRoot.Descendants<Cell>().Where(cell => cell.CellFormula != null)) {
                    string? dependencyReference = NormalizeFormulaCellReference(dependencyCell.CellReference?.Value);
                    if (dependencyReference == null
                        || !TryParseCellReference(dependencyReference, out int dependencyRow, out int dependencyColumn)
                        || dependencyRow < r1 || dependencyRow > r2 || dependencyColumn < c1 || dependencyColumn > c2) {
                        continue;
                    }

                    string formattedDependencyCell = $"{dependencySheet.Name}!{dependencyReference}";
                    if (string.Equals(dependencySheet.Name, Name, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(dependencyReference, sourceReference, StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    if (dependencyCell.CellValue == null) {
                        issues.Add($"Dependency '{formattedDependencyCell}' is a formula without a cached result.");
                    }

                    string dependencyFormula = dependencyCell.CellFormula!.Text ?? string.Empty;
                    if (!TryEvaluateFormulaValue(dependencyFormula, out _)) {
                        issues.Add($"Dependency '{formattedDependencyCell}' contains a formula outside the lightweight evaluator support.");
                    }
                }
            }

            return issues
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(issue => issue, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private string NormalizeFormulaDependencyReference(string reference) {
            string normalized = reference.Trim().Replace("$", string.Empty);
            if (TryResolveFormulaRangeReference(normalized, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2)) {
                string start = A1.CellReference(r1, c1);
                string end = A1.CellReference(r2, c2);
                return r1 == r2 && c1 == c2
                    ? $"{sheet.Name}!{start}"
                    : $"{sheet.Name}!{start}:{end}";
            }

            return normalized;
        }

        private static string MaskFormulaStringLiterals(string formula) {
            var builder = new StringBuilder(formula.Length);
            bool inString = false;
            for (int i = 0; i < formula.Length; i++) {
                char character = formula[i];
                if (character == '"') {
                    builder.Append(' ');
                    if (inString && i + 1 < formula.Length && formula[i + 1] == '"') {
                        i++;
                        builder.Append(' ');
                        continue;
                    }

                    inString = !inString;
                    continue;
                }

                builder.Append(inString ? ' ' : character);
            }

            return builder.ToString();
        }

        private static bool TryParseCellReference(string reference, out int row, out int column) {
            row = 0;
            column = 0;
            try {
                var cellRef = A1.ParseCellRef(reference.Replace("$", string.Empty));
                row = cellRef.Row;
                column = cellRef.Col;
                return row > 0 && column > 0 && row <= A1.MaxRows && column <= A1.MaxColumns;
            } catch (ArgumentException) {
                return false;
            }
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
            return value.ErrorCode ?? value.Text ?? value.Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private bool TryResolveFormulaArgument(string token, out FormulaArgumentValue value) {
            string trimmed = token.Trim();
            if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"') {
                value = new FormulaArgumentValue(null, trimmed.Substring(1, trimmed.Length - 2).Replace("\"\"", "\""));
                return true;
            }

            if (trimmed.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) {
                value = new FormulaArgumentValue(1d, "1");
                return true;
            }

            if (trimmed.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) {
                value = new FormulaArgumentValue(0d, "0");
                return true;
            }

            if (TryParseFormulaErrorLiteral(trimmed, out string errorCode)) {
                value = FormulaArgumentValue.Error(errorCode);
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

                if (TryResolveFormulaArgument(trimmed, out var argumentValue)) {
                    values.Add(argumentValue);
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

        private static bool TryConvertFormulaAValues(IReadOnlyList<FormulaArgumentValue> values, out List<double> numbers) {
            numbers = new List<double>();
            foreach (var value in values) {
                if (value.IsUnresolvedFormula || value.IsError) {
                    numbers.Clear();
                    return false;
                }

                if (value.Number.HasValue) {
                    numbers.Add(value.Number.Value);
                    continue;
                }

                if (value.Text != null) {
                    numbers.Add(0d);
                }
            }

            return true;
        }

        private bool TryResolveNumericOperand(string token, out double value) {
            token = token.Trim();
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

        private bool TryResolveFormulaReferenceArgument(string token, out ExcelSheet sheet, out int row, out int column) {
            if (TryParseQualifiedFormulaCellReference(token, out sheet, out row, out column)) {
                return true;
            }

            if (TryResolveFormulaRangeReference(token, out sheet, out row, out column, out int endRow, out int endColumn)
                && row == endRow
                && column == endColumn) {
                return true;
            }

            sheet = this;
            row = 0;
            column = 0;
            return false;
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

            if (value.Kind == ExcelCellDataKind.Error) {
                return FormulaArgumentValue.Error(value.CachedText ?? value.Value?.ToString() ?? "#VALUE!");
            }

            if (TryParseFormulaErrorLiteral(value.CachedText ?? value.Value?.ToString() ?? string.Empty, out string errorCode)) {
                return FormulaArgumentValue.Error(errorCode);
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

        private static bool TryParseFormulaErrorLiteral(string token, out string errorCode) {
            string value = token.Trim();
            if (value.StartsWith("=", StringComparison.Ordinal)) {
                value = value.Substring(1).Trim();
            }

            switch (value.ToUpperInvariant()) {
                case "#NULL!":
                    errorCode = "#NULL!";
                    return true;
                case "#DIV/0!":
                    errorCode = "#DIV/0!";
                    return true;
                case "#VALUE!":
                    errorCode = "#VALUE!";
                    return true;
                case "#REF!":
                    errorCode = "#REF!";
                    return true;
                case "#NAME?":
                    errorCode = "#NAME?";
                    return true;
                case "#NUM!":
                    errorCode = "#NUM!";
                    return true;
                case "#N/A":
                    errorCode = "#N/A";
                    return true;
                default:
                    errorCode = string.Empty;
                    return false;
            }
        }

        private readonly struct FormulaArgumentValue {
            internal FormulaArgumentValue(double? number, string? text, bool isUnresolvedFormula = false, bool isError = false) {
                Number = number;
                Text = text;
                IsUnresolvedFormula = isUnresolvedFormula;
                IsError = isError;
            }

            internal double? Number { get; }
            internal string? Text { get; }
            internal bool IsUnresolvedFormula { get; }
            internal bool IsError { get; }
            internal string? ErrorCode => IsError ? Text : null;
            internal bool HasValue => Number.HasValue || Text != null || IsError;

            internal static FormulaArgumentValue UnresolvedFormula() {
                return new FormulaArgumentValue(null, null, isUnresolvedFormula: true);
            }

            internal static FormulaArgumentValue Error(string errorCode) {
                return new FormulaArgumentValue(null, errorCode, isError: true);
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
