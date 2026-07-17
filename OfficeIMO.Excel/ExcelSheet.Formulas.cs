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
        private Dictionary<string, int>? _formulaEvaluationDepthCache;
        private HashSet<string>? _formulaEvaluationStack;
        private Stack<FormulaEvaluationDepthFrame>? _formulaEvaluationDepthFrames;
        private FormulaEvaluationGuardState? _formulaEvaluationGuardState;
        private string? _formulaEvaluationCellReference;

        private sealed class FormulaEvaluationGuardState {
            internal bool DependencyGuardBlocked { get; set; }
        }

        private sealed class FormulaEvaluationDepthFrame {
            internal int MaximumChildDepth { get; private set; }
            internal bool DependencyGuardBlocked { get; private set; }

            internal void IncludeChild(int depth) {
                if (depth > MaximumChildDepth) {
                    MaximumChildDepth = depth;
                }
            }

            internal void BlockByDependencyGuard() {
                DependencyGuardBlocked = true;
            }
        }

        private static readonly Regex SimpleFunctionFormulaRegex = new Regex(
            @"^\s*=?\s*(SUM|AVERAGE|AVERAGEA|MIN|MINA|MAX|MAXA|COUNT|COUNTA|COUNTBLANK|SUBTOTAL|COUNTIF|SUMIF|AVERAGEIF|COUNTIFS|SUMIFS|AVERAGEIFS|MINIFS|MAXIFS|PRODUCT|MEDIAN|LARGE|SMALL|MODE\.SNGL|MODE|GEOMEAN|HARMEAN|AVEDEV|DEVSQ|SUMXMY2|SUMX2MY2|SUMX2PY2|SUMSQ|SUMPRODUCT|STDEV\.S|STDEV\.P|VAR\.S|VAR\.P|PERCENTILE\.INC|PERCENTILE\.EXC|QUARTILE\.INC|QUARTILE\.EXC|PERCENTRANK\.INC|PERCENTRANK\.EXC|RANK\.EQ|RANK\.AVG|COVAR|COVARIANCE\.P|COVARIANCE\.S|CORREL|SLOPE|INTERCEPT|RSQ|FORECAST\.LINEAR|PMT|PV|FV|NPER|NPV|VLOOKUP|HLOOKUP|XLOOKUP|INDEX|MATCH|XMATCH|ABS|SIGN|ROUND|ROUNDUP|ROUNDDOWN|MROUND|TRUNC|INT|CEILING\.MATH|FLOOR\.MATH|CEILING|FLOOR|POWER|SQRT|LN|LOG10|EXP|PI|RADIANS|DEGREES|MOD|ROW|COLUMN|ROWS|COLUMNS|DATE|TIME|DATEVALUE|TIMEVALUE|TODAY|NOW|YEAR|MONTH|DAY|HOUR|MINUTE|SECOND|DATEDIF|YEARFRAC|EDATE|EOMONTH|DAYS|DAYS360|WEEKDAY|WEEKNUM|ISOWEEKNUM|NETWORKDAYS|WORKDAY\.INTL|WORKDAY|IF|IFS|SWITCH|CHOOSE|ISBLANK|ISNUMBER|ISTEXT|ISERROR|ISERR|ISNA|ISFORMULA|AND|OR|NOT|IFERROR|IFNA|CONCAT|CONCATENATE|TEXT|TEXTJOIN|TEXTBEFORE|TEXTAFTER|FORMULATEXT|LEFT|RIGHT|MID|LEN|TRIM|UPPER|LOWER|PROPER|SUBSTITUTE|FIND|SEARCH|VALUE|EXACT|REPT)\s*\((.*)\)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        private static readonly Regex FunctionNameFormulaRegex = new Regex(
            @"^\s*=?\s*([A-Za-z_][A-Za-z0-9_.]*)\s*\(",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        private static readonly Regex AnyFunctionFormulaRegex = new Regex(
            @"^\s*=?\s*([A-Za-z_][A-Za-z0-9_.]*)\s*\((.*)\)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            FormulaRegexTimeout);

        private static readonly Regex FormulaReferenceRegex = new Regex(
            @"(?<![A-Za-z0-9_\.])(?<reference>(?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_ .]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![A-Za-z0-9_\.]|\s*\()",
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
            MaterializePendingDirectCellValues();

            int count = 0;
            WriteLock(() => {
                MaterializePendingDirectCellValues();

                var previousCache = _formulaEvaluationCache;
                var previousDepthCache = _formulaEvaluationDepthCache;
                var previousStack = _formulaEvaluationStack;
                var previousDepthFrames = _formulaEvaluationDepthFrames;
                var previousGuardState = _formulaEvaluationGuardState;
                _formulaEvaluationCache = new Dictionary<string, FormulaArgumentValue>(StringComparer.OrdinalIgnoreCase);
                _formulaEvaluationDepthCache = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                _formulaEvaluationStack = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _formulaEvaluationDepthFrames = new Stack<FormulaEvaluationDepthFrame>();
                _formulaEvaluationGuardState = new FormulaEvaluationGuardState();
                bool changed = false;

                try {
                    foreach (var cell in WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula != null).ToList()) {
                        _formulaEvaluationGuardState.DependencyGuardBlocked = false;
                        if (!TryEvaluateFormulaCellValue(cell, out FormulaArgumentValue result)) {
                            if (_formulaEvaluationGuardState.DependencyGuardBlocked) {
                                if (cell.CellValue != null) {
                                    cell.CellValue = null;
                                    changed = true;
                                }

                                if (cell.CellFormula!.CalculateCell?.Value != true) {
                                    cell.CellFormula.CalculateCell = true;
                                    changed = true;
                                }
                            }

                            continue;
                        }

                        SetFormulaCachedValue(cell, result);
                        cell.CellFormula!.CalculateCell = false;
                        changed = true;
                        count++;
                    }
                } finally {
                    _formulaEvaluationCache = previousCache;
                    _formulaEvaluationDepthCache = previousDepthCache;
                    _formulaEvaluationStack = previousStack;
                    _formulaEvaluationDepthFrames = previousDepthFrames;
                    _formulaEvaluationGuardState = previousGuardState;
                }

                if (changed) {
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
            string? previousCellReference = _formulaEvaluationCellReference;
            _formulaEvaluationCellReference = reference;
            try {
                if (reference == null
                    || _formulaEvaluationCache == null
                    || _formulaEvaluationDepthCache == null
                    || _formulaEvaluationStack == null
                    || _formulaEvaluationDepthFrames == null) {
                    return TryEvaluateFormulaValue(cell.CellFormula.Text ?? string.Empty, out result);
                }

                string cacheKey = GetFormulaEvaluationCacheKey(reference);
                if (_formulaEvaluationCache.TryGetValue(cacheKey, out FormulaArgumentValue cachedResult)) {
                    if (!_formulaEvaluationDepthCache.TryGetValue(cacheKey, out int cachedDepth)
                        || _formulaEvaluationStack.Count + cachedDepth > _excelDocument.Calculation.MaximumDependencyDepth) {
                        BlockCurrentFormulaByDependencyGuard();
                        return false;
                    }

                    if (_formulaEvaluationDepthFrames.Count > 0) {
                        _formulaEvaluationDepthFrames.Peek().IncludeChild(cachedDepth);
                    }

                    result = cachedResult;
                    return true;
                }

                if (_formulaEvaluationStack.Count >= _excelDocument.Calculation.MaximumDependencyDepth) {
                    BlockCurrentFormulaByDependencyGuard();
                    return false;
                }

                if (!_formulaEvaluationStack.Add(cacheKey)) {
                    BlockCurrentFormulaByDependencyGuard();
                    return false;
                }

                var depthFrame = new FormulaEvaluationDepthFrame();
                _formulaEvaluationDepthFrames.Push(depthFrame);
                bool evaluated = false;
                int evaluationDepth = 0;
                try {
                    if (!TryEvaluateFormulaValue(cell.CellFormula.Text ?? string.Empty, out result)) {
                        return false;
                    }

                    evaluationDepth = depthFrame.MaximumChildDepth + 1;
                    _formulaEvaluationCache[cacheKey] = result;
                    _formulaEvaluationDepthCache[cacheKey] = evaluationDepth;
                    evaluated = true;
                    return true;
                } finally {
                    _formulaEvaluationDepthFrames.Pop();
                    _formulaEvaluationStack.Remove(cacheKey);
                    if (evaluated && _formulaEvaluationDepthFrames.Count > 0) {
                        _formulaEvaluationDepthFrames.Peek().IncludeChild(evaluationDepth);
                    } else if (depthFrame.DependencyGuardBlocked && _formulaEvaluationDepthFrames.Count > 0) {
                        _formulaEvaluationDepthFrames.Peek().BlockByDependencyGuard();
                    }
                }
            } finally {
                _formulaEvaluationCellReference = previousCellReference;
            }
        }

        private void BlockCurrentFormulaByDependencyGuard() {
            if (_formulaEvaluationGuardState != null) {
                _formulaEvaluationGuardState.DependencyGuardBlocked = true;
            }

            if (_formulaEvaluationDepthFrames != null && _formulaEvaluationDepthFrames.Count > 0) {
                _formulaEvaluationDepthFrames.Peek().BlockByDependencyGuard();
            }
        }

        private static void SetFormulaCachedValue(Cell cell, FormulaArgumentValue result) {
            if (result.Number.HasValue) {
                cell.CellValue = new CellValue(InvariantNumberText.Get(result.Number.Value));
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
                FormulaDependencyAliasCatalog dependencyAliases = GetFormulaDependencyAliases();
                foreach (var cell in WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula != null)) {
                    string formula = cell.CellFormula!.Text ?? string.Empty;
                    bool supported = TryEvaluateFormulaCellValue(cell, out _);
                    IReadOnlyList<string> dependencies = GetFormulaDependencies(formula, dependencyAliases);
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

        internal void SetLegacyArrayFormula(string a1Range, string formula) {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentNullException(nameof(formula));
            if (!A1.TryParseRange(a1Range, out int r1, out int c1, out _, out _)) {
                (r1, c1) = A1.ParseCellRef(a1Range);
                if (r1 <= 0 || c1 <= 0) {
                    throw new ArgumentException($"Invalid A1 range '{a1Range}'.", nameof(a1Range));
                }
            }

            WriteLock(() => {
                var topLeft = GetCell(r1, c1);
                topLeft.CellFormula = new CellFormula(Utilities.ExcelSanitizer.SanitizeFormula(formula)) {
                    FormulaType = CellFormulaValues.Array,
                    Reference = a1Range
                };
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

            formula = NormalizeSupportedFunctionPrefix(formula);
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

                if (!functionMatch.Success && TryEvaluateCustomFormulaFunction(formula, out result)) {
                    return true;
                }
            } catch (RegexMatchTimeoutException) {
                return false;
            }

            if (TryEvaluateFormula(formula, out double numeric)) {
                result = new FormulaArgumentValue(numeric, InvariantNumberText.Get(numeric));
                return true;
            }

            return false;
        }

        private bool TryEvaluateFormula(string formula, out double result) {
            result = 0;
            if (string.IsNullOrWhiteSpace(formula) || formula.Length > MaxSupportedFormulaLength) {
                return false;
            }

            formula = NormalizeSupportedFunctionPrefix(formula);
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

                if (TryEvaluateCustomFormulaFunction(formula, out FormulaArgumentValue customResult)
                    && customResult.Number.HasValue) {
                    result = customResult.Number.Value;
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

    }
}
