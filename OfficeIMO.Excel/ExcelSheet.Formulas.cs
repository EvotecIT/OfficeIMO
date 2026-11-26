using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int MaxSupportedFormulaLength = 8192;
        private FormulaEvaluationResults? _formulaEvaluationCache;
        private Dictionary<long, int>? _formulaEvaluationDepthCache;
        private HashSet<long>? _formulaEvaluationStack;
        private Stack<FormulaEvaluationDepthFrame>? _formulaEvaluationDepthFrames;
        private FormulaEvaluationGuardState? _formulaEvaluationGuardState;
        private IReadOnlyDictionary<uint, SharedFormulaDefinition>? _formulaEvaluationSharedDefinitions;
        private Dictionary<string, IReadOnlyDictionary<uint, SharedFormulaDefinition>>? _formulaEvaluationSharedDefinitionsBySheet;
        private string? _formulaEvaluationCellReference;

        internal sealed class FormulaEvaluationGuardState {
            internal bool DependencyGuardBlocked { get; set; }
            internal HashSet<string> ActiveDefinedNames { get; } = new(StringComparer.OrdinalIgnoreCase);
        }

        internal sealed class FormulaEvaluationDepthFrame {
            private const int MaximumPooledFrames = 64;
            [ThreadStatic]
            private static Stack<FormulaEvaluationDepthFrame>? t_pool;

            /// <summary>Returns a cleared frame; a calculation pass needs one per formula cell it evaluates.</summary>
            internal static FormulaEvaluationDepthFrame Rent() {
                Stack<FormulaEvaluationDepthFrame>? pool = t_pool;
                if (pool == null || pool.Count == 0) return new FormulaEvaluationDepthFrame();
                FormulaEvaluationDepthFrame frame = pool.Pop();
                frame.MaximumChildDepth = 0;
                frame.DependencyGuardBlocked = false;
                frame.UsedUnevaluatedFormulaCache = false;
                return frame;
            }

            /// <summary>Takes back a frame that is off the evaluation stack and no longer read.</summary>
            internal static void Return(FormulaEvaluationDepthFrame frame) {
                Stack<FormulaEvaluationDepthFrame> pool = t_pool ??= new Stack<FormulaEvaluationDepthFrame>();
                if (pool.Count < MaximumPooledFrames) pool.Push(frame);
            }

            internal int MaximumChildDepth { get; private set; }
            internal bool DependencyGuardBlocked { get; private set; }
            internal bool UsedUnevaluatedFormulaCache { get; private set; }

            internal void IncludeChild(int depth) {
                if (depth > MaximumChildDepth) {
                    MaximumChildDepth = depth;
                }
            }

            internal void BlockByDependencyGuard() {
                DependencyGuardBlocked = true;
            }

            internal void MarkUnevaluatedFormulaCache() {
                UsedUnevaluatedFormulaCache = true;
            }
        }

        /// <summary>
        /// Evaluates supported formulas on this sheet and writes cached results.
        /// </summary>
        public int RecalculateSupportedFormulas() {
            using var context = new FormulaCalculationContext();
            return RecalculateSupportedFormulas(context);
        }

        private int _formulaCellCountHint;

        internal int RecalculateSupportedFormulas(FormulaCalculationContext context) {
            MaterializePendingDirectCellValues();

            int count = 0;
            WriteLock(() => {
                MaterializePendingDirectCellValues();
                // Formulas reach other sheets too, whose buffered cells must be in their worksheets by then.
                MaterializeDeferredDataSetImportIfNeeded();
                // The write lock makes this thread the workbook's only evaluator, so the pass may reuse parsed syntax.
                using ExcelFormulaSyntaxCache.Scope syntaxScope = _excelDocument.FormulaSyntaxCache.Enter();
                long formulaInputMutationVersion = _excelDocument.CaptureFormulaInputMutationVersion();

                var previousCache = _formulaEvaluationCache;
                var previousCalculationContext = _formulaCalculationContext;
                var previousDepthCache = _formulaEvaluationDepthCache;
                var previousStack = _formulaEvaluationStack;
                var previousDepthFrames = _formulaEvaluationDepthFrames;
                var previousGuardState = _formulaEvaluationGuardState;
                var previousSharedDefinitions = _formulaEvaluationSharedDefinitions;
                var previousSharedDefinitionsBySheet = _formulaEvaluationSharedDefinitionsBySheet;
                _formulaEvaluationCache = context.Cache;
                _formulaCalculationContext = context;
                _formulaEvaluationDepthCache = context.DepthCache;
                _formulaEvaluationStack = context.Stack;
                _formulaEvaluationDepthFrames = context.DepthFrames;
                _formulaEvaluationGuardState = context.GuardState;
                bool changed = false;
                bool allFormulasEvaluated = true;
                bool passCompleted = false;
                var recalculatedCells = new List<string>();

                try {
                    // One walk lists the formula cells for the pass, the shared-formula check and the array index.
                    // The previous pass's count sizes the list, so a sheet recalculated again copies no growing buffer.
                    // Inside the try, so a failure here still restores the sheet's calculation state.
                    var formulaCells = new List<ExcelFormulaCell>(_formulaCellCountHint);
                    foreach (ExcelFormulaCell formulaCell in EnumerateFormulaCells()) formulaCells.Add(formulaCell);
                    _formulaCellCountHint = formulaCells.Count;
                    GetFixedArraySheetIndex(formulaCells);
                    _formulaEvaluationSharedDefinitions = HasSharedFormulas(formulaCells)
                        ? BuildSharedFormulaDefinitions()
                        : new Dictionary<uint, SharedFormulaDefinition>();
                    _formulaEvaluationSharedDefinitionsBySheet = new Dictionary<string, IReadOnlyDictionary<uint, SharedFormulaDefinition>>(
                        StringComparer.OrdinalIgnoreCase) {
                        [Name] = _formulaEvaluationSharedDefinitions
                    };
                    if (_excelDocument.HasCellMetadataPart) PlanDynamicArrayOwners();
                    context.ReserveFormulaCells(Name, formulaCells.Count);
                    recalculatedCells.Capacity = formulaCells.Count;
                    foreach (var cell in formulaCells) {
                        _formulaEvaluationGuardState.DependencyGuardBlocked = false;
                        if (!TryEvaluateFormulaCellValue(cell, out FormulaArgumentValue result)) {
                            allFormulasEvaluated = false;
                            if (_formulaEvaluationGuardState.DependencyGuardBlocked) {
                                if (ClearFormulaCachedValue(cell)) changed = true;

                                if (cell.CalculateCell != true) {
                                    cell.SetCalculateCell(true);
                                    changed = true;
                                }
                            }

                            continue;
                        }

                        if (TryGetDynamicArrayPlan(cell, out DynamicArrayPlan dynamicPlan)) {
                            WriteDynamicArrayFormulaCache(dynamicPlan);
                        } else {
                            SetFormulaCachedValue(cell, result);
                            if (TryGetCalculatedArray(cell, out FormulaArrayValue array)) WriteFixedArrayFormulaCache(cell, array);
                        }
                        // Writing an equal value would only replace the attribute object.
                        if (cell.CalculateCell != false) cell.SetCalculateCell(false);
                        recalculatedCells.Add(cell.Reference ?? string.Empty);
                        changed = true;
                        count++;
                    }
                    passCompleted = true;
                } finally {
                    // A sheet whose every formula was evaluated is marked as a whole below, which answers each of its
                    // cells as their own marks would; otherwise, or when the pass stops early, each cell keeps its mark.
                    if (!passCompleted || !allFormulasEvaluated) {
                        foreach (string reference in recalculatedCells) {
                            MarkFormulaCellRecalculated(reference, formulaInputMutationVersion);
                        }
                    }
                    _formulaEvaluationCache = previousCache;
                    _formulaCalculationContext = previousCalculationContext;
                    _formulaEvaluationDepthCache = previousDepthCache;
                    _formulaEvaluationStack = previousStack;
                    _formulaEvaluationDepthFrames = previousDepthFrames;
                    _formulaEvaluationGuardState = previousGuardState;
                    _formulaEvaluationSharedDefinitions = previousSharedDefinitions;
                    _formulaEvaluationSharedDefinitionsBySheet = previousSharedDefinitionsBySheet;
                }

                if (changed) {
                    _hasWorksheetMutations = true;
                    MarkRequiresSavePreparation();
                    ClearCellTextSharedStringCache();
                }

                if (allFormulasEvaluated) {
                    MarkFormulaSheetRecalculated(formulaInputMutationVersion);
                }

            });

            return count;
        }

        private bool TryEvaluateFormulaCell(ExcelFormulaCell cell, out double result) {
            result = 0;
            if (!TryEvaluateFormulaCellValue(cell, out FormulaArgumentValue value) || !value.Number.HasValue) {
                return false;
            }

            result = value.Number.Value;
            return true;
        }

        private bool TryEvaluateFormulaCellValue(
            ExcelFormulaCell cell,
            out FormulaArgumentValue result,
            IReadOnlyDictionary<uint, SharedFormulaDefinition>? sharedFormulaDefinitions = null) {
            result = default;
            if (!cell.HasFormula) {
                return false;
            }
            // Imported formulas must not recalculate schema-owned literal column names.
            if (FormulaTargetsTableHeader(cell)) return false;

            string formula = ResolveCellFormulaText(
                cell,
                sharedFormulaDefinitions ?? _formulaEvaluationSharedDefinitions);

            string? reference = NormalizeFormulaCellReference(cell.Reference);
            string? previousCellReference = _formulaEvaluationCellReference;
            // Expression nesting belongs to a single formula. Dependency cells have
            // their own nesting budget and are bounded by MaximumDependencyDepth.
            int previousScalarDepth = _scalarFormulaEvaluationDepth;
            _formulaEvaluationCellReference = reference;
            _scalarFormulaEvaluationDepth = 0;
            try {
                if (reference == null
                    || _formulaEvaluationCache == null
                    || _formulaEvaluationDepthCache == null
                    || _formulaEvaluationStack == null
                    || _formulaEvaluationDepthFrames == null) {
                    bool supported = cell.FormulaKind == ExcelFormulaKind.Array
                        ? TryEvaluateAuthoredFormulaCell(cell, formula, out result)
                        : TryEvaluateFormulaValue(formula, out result);
                    if (supported) result = NormalizeLookupResult(result);
                    return supported;
                }

                long cacheKey = GetFormulaEvaluationCacheKey(reference);
                if (_formulaEvaluationCache.TryGetValue(cacheKey, out FormulaArgumentValue cachedResult)) {
                    if (!_formulaEvaluationDepthCache.TryGetValue(cacheKey, out int cachedDepth)
                        || _formulaEvaluationStack.Count + cachedDepth > _excelDocument.Calculation.MaximumDependencyDepth) {
                        BlockCurrentFormulaByDependencyGuard();
                        return false;
                    }

                    if (_formulaEvaluationDepthFrames.Count > 0) {
                        _formulaEvaluationDepthFrames.Peek().IncludeChild(cachedDepth);
                        if (cachedResult.IsUnevaluatedFormulaCache)
                            _formulaEvaluationDepthFrames.Peek().MarkUnevaluatedFormulaCache();
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

                var depthFrame = FormulaEvaluationDepthFrame.Rent();
                _formulaEvaluationDepthFrames.Push(depthFrame);
                bool evaluated = false;
                int evaluationDepth = 0;
                try {
                    if (!(cell.FormulaKind == ExcelFormulaKind.Array
                        ? TryEvaluateAuthoredFormulaCell(cell, formula, out result)
                        : TryEvaluateFormulaValue(formula, out result))) {
                        return false;
                    }
                    if (depthFrame.DependencyGuardBlocked) {
                        return false;
                    }

                    // A blank reference remains blank inside an expression. A formula
                    // cell stores its value, so dependants see the displayed numeric zero.
                    result = NormalizeLookupResult(result);
                    if (depthFrame.UsedUnevaluatedFormulaCache)
                        result = result.WithUnevaluatedFormulaCache();

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
                        if (result.IsUnevaluatedFormulaCache)
                            _formulaEvaluationDepthFrames.Peek().MarkUnevaluatedFormulaCache();
                    } else if (depthFrame.DependencyGuardBlocked && _formulaEvaluationDepthFrames.Count > 0) {
                        _formulaEvaluationDepthFrames.Peek().BlockByDependencyGuard();
                    }
                    FormulaEvaluationDepthFrame.Return(depthFrame);
                }
            } finally {
                _formulaEvaluationCellReference = previousCellReference;
                _scalarFormulaEvaluationDepth = previousScalarDepth;
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

        private static string? GetFormulaCachedText(ExcelFormulaCell cell) {
            string? text = cell.ValueText;
            return text != null && cell.ValueType == ExcelCellValueType.String
                ? Utilities.ExcelSanitizer.DecodeXString(text) : text;
        }

        /// <summary>
        /// Returns the formula text from a cell, if present.
        /// </summary>
        public string? GetFormulaText(int row, int column) {
            ExcelFormulaCell cell = GetFormulaCell(row, column);
            return !cell.HasFormula ? null : ResolveCellFormulaText(cell);
        }

        /// <summary>
        /// Tries to return a formula cell's cached value, decoding string escapes while retaining numeric text precision.
        /// </summary>
        public bool TryGetCachedFormulaValue(int row, int column, out string? value) {
            // A formula cell is always a DOM cell, so a stored cell stays stored.
            ExcelFormulaCell cell = GetFormulaCell(row, column);
            value = !cell.HasFormula ? null : ResolveRichValueError(cell, GetFormulaCachedText(cell));
            return value != null;
        }

        private bool HasSufficientFormulaExecutionStack() {
            try {
                System.Runtime.CompilerServices.RuntimeHelpers.EnsureSufficientExecutionStack();
                return true;
            } catch (InsufficientExecutionStackException) {
                // Independent cell and expression limits must also fit the caller's stack.
                BlockCurrentFormulaByDependencyGuard();
                return false;
            }
        }

        private int _scalarFormulaEvaluationDepth;

        private bool TryEvaluateFormulaValueCore(string formula, out FormulaArgumentValue result) {
            result = default;
            if (string.IsNullOrWhiteSpace(formula) || formula.Length > MaxSupportedFormulaLength) {
                return false;
            }

            formula = NormalizeSupportedFunctionPrefix(formula);
            string literal = formula.Trim().TrimStart('=').Trim();
            if (literal.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || literal.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) {
                result = FormulaArgumentValue.Boolean(literal.Equals("TRUE", StringComparison.OrdinalIgnoreCase));
                return true;
            }
            ExcelFormulaExpressionParser.TryParseSupportedFunctionCall(formula, out ExcelFormulaFunctionCallSyntax? functionCall);
            if (functionCall != null) {
                    string function = functionCall.UpperName;
                    string args = functionCall.Arguments;
                    if (function == "MATCH" || function == "XMATCH") return TryEvaluateMatchValue(function, args, out result);
                    if (function is "INT" or "MOD" or "SQRT" or "SIGN" or "TRUNC" or "ROUNDUP" or "ROUNDDOWN" or "EXP" or "LN" or "LOG" or "LOG10" or "CEILING" or "FLOOR" or "EVEN" or "ODD" or "FACT" or "COMBIN") return TryEvaluateScalarMathValue(function, args, out result);
                    if (function is "GCD" or "LCM") return TryEvaluateIntegerAggregateValue(function, args, out result);
                    if (function == "PROB") return TryEvaluateProbabilityValue(args, out result);
                    if (function == "SUMPRODUCT") return TryEvaluateSumProductValue(args, out result);
                    if (function is "IPMT" or "PPMT") return TryEvaluatePaymentSplitValue(function, args, out result);
                    if (function is "RATE" or "IRR" or "XIRR" or "XNPV") return TryEvaluateCashFlowValue(function, args, out result);
                    if (function == "RANDBETWEEN") return TryEvaluateRandomBetweenValue(args, out result);
                    if (function == "OFFSET") return TryEvaluateOffsetValue(args, out result);
                    if (function == "INDIRECT") return TryEvaluateIndirectValue(args, out result);
                    if (function == "LOOKUP") return TryEvaluateLookupFunctionValue(args, out result);
                    if (function == "ADDRESS") return TryEvaluateAddressValue(args, out result);
                    if (IsDistributionFunction(function)) return TryEvaluateDistributionValue(function, args, out result);
                    if (function == "GETPIVOTDATA") return TryEvaluatePivotDataValue(args, out result);
                    if ((function == "TRUE" || function == "FALSE") && string.IsNullOrWhiteSpace(args)) {
                        bool boolean = function == "TRUE";
                        result = new FormulaArgumentValue(boolean ? 1 : 0, function, isBoolean: true);
                        return true;
                    }
                    if (function == "NA" && string.IsNullOrWhiteSpace(args)) {
                        result = FormulaArgumentValue.Error("#N/A");
                        return true;
                    }
                    if (function == "DATE") return TryEvaluateDateValue(args, out result);
                    if (function == "DATEDIF") return TryEvaluateDateDifValue(args, out result);
                    if (function == "IFERROR") return TryEvaluateIfErrorValue(args, out result);
                    if (function == "IFNA") return TryEvaluateIfNaValue(args, out result);

                    if (function == "IF") return TryEvaluateIfValue(args, out result);
                    if (function == "IFS") return TryEvaluateIfsValue(args, out result);
                    if (function == "SWITCH") return TryEvaluateSwitchValue(args, out result);
                    if (function == "CHOOSE") return TryEvaluateChooseValue(args, out result);

                    if (function == "ISBLANK" || function == "ISNUMBER" || function == "ISLOGICAL" || function == "ISTEXT"
                        || function == "ISERROR" || function == "ISERR" || function == "ISNA" || function == "ISFORMULA")
                        return TryEvaluateInfoFunction(function, args, out result);

                    if (function == "AVERAGEA" || function == "MINA" || function == "MAXA")
                        return TryEvaluateAValueAggregate(function, args, out result);

                    if (IsSupportedTextFunction(function)) return TryEvaluateTextFunction(function, args, out result);

                    if (function == "VLOOKUP" || function == "HLOOKUP" || function == "XLOOKUP")
                        return TryEvaluateLookupValue(function, args, out result);

                    if (function == "INDEX") return TryEvaluateIndexValue(args, out result);
                }

            if (functionCall == null && TryEvaluateCustomFormulaFunction(formula, out result)) {
                return true;
            }

            // INDEX and OFFSET, the only calls that yield a reference, returned above.
            if (functionCall == null && TryEvaluateSingleReferenceFormulaValue(formula, out result)) {
                return true;
            }

            if (TryEvaluateFormulaCore(formula, out double numeric, out FormulaArgumentValue error)) {
                bool isBoolean = functionCall != null && (functionCall.Name.Equals("AND", StringComparison.OrdinalIgnoreCase)
                    || functionCall.Name.Equals("OR", StringComparison.OrdinalIgnoreCase) || functionCall.Name.Equals("NOT", StringComparison.OrdinalIgnoreCase));
                result = FormulaArgumentValue.FromNumber(numeric, isBoolean);
                return true;
            }

            if (error.IsError) { result = error; return true; }

            return false;
        }

        private bool TryEvaluateFormula(string formula, out double result) {
            return TryEvaluateFormulaCore(formula, out result, out _);
        }

        private bool TryEvaluateFormulaCore(string formula, out double result, out FormulaArgumentValue error) {
            result = 0;
            error = default;
            if (string.IsNullOrWhiteSpace(formula) || formula.Length > MaxSupportedFormulaLength) {
                return false;
            }

            formula = NormalizeSupportedFunctionPrefix(formula);
            ExcelFormulaExpressionParser.TryParseSupportedFunctionCall(formula, out ExcelFormulaFunctionCallSyntax? functionCall);
            if (functionCall != null) {
                    string function = functionCall.UpperName;
                    string args = functionCall.Arguments;
                    if (IsRangeAggregateFunction(function)) {
                        bool aggregated = TryEvaluateRangeAggregate(function, args, out result, out error, out bool handled);
                        if (handled) return aggregated;
                    }
                    if (function is "INT" or "MOD" or "SQRT" or "SIGN" or "TRUNC" or "ROUNDUP" or "ROUNDDOWN" or "EXP" or "LN" or "LOG" or "LOG10" or "CEILING" or "FLOOR" or "EVEN" or "ODD" or "FACT" or "COMBIN") {
                        if (!TryEvaluateScalarMathValue(function, args, out FormulaArgumentValue value)) return false;
                        if (value.IsError) { error = value; return false; }
                        result = value.Number!.Value;
                        return true;
                    }
                    if (function is "GCD" or "LCM" or "PROB" or "RANDBETWEEN" or "OFFSET" or "INDIRECT" or "LOOKUP" || IsDistributionFunction(function)) {
                        bool evaluated = function is "GCD" or "LCM" ? TryEvaluateIntegerAggregateValue(function, args, out FormulaArgumentValue value)
                            : function == "PROB" ? TryEvaluateProbabilityValue(args, out value)
                            : function == "RANDBETWEEN" ? TryEvaluateRandomBetweenValue(args, out value)
                            : function == "INDIRECT" ? TryEvaluateIndirectValue(args, out value)
                            : function == "LOOKUP" ? TryEvaluateLookupFunctionValue(args, out value)
                            : function == "OFFSET" ? TryEvaluateOffsetValue(args, out value)
                            : TryEvaluateDistributionValue(function, args, out value);
                        if (!evaluated) return false;
                        if (value.IsError) { error = value; return false; }
                        if (!value.Number.HasValue) return false;
                        result = value.Number.Value;
                        return true;
                    }
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

                    if (function == "ISBLANK" || function == "ISNUMBER" || function == "ISLOGICAL" || function == "ISTEXT" || function == "ISERROR" || function == "ISERR" || function == "ISNA" || function == "ISFORMULA") {
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
                        return TryEvaluateConditionalAggregate(function, args, out result, out error);
                    }

                    if (function == "COUNTIFS" || function == "SUMIFS" || function == "AVERAGEIFS" || function == "MINIFS" || function == "MAXIFS") {
                        return TryEvaluateMultiCriteriaAggregate(function, args, out result, out error);
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

                    if (function is "IPMT" or "PPMT" or "RATE" or "IRR" or "XIRR" or "XNPV") {
                        bool evaluated = function is "IPMT" or "PPMT"
                            ? TryEvaluatePaymentSplitValue(function, args, out FormulaArgumentValue paymentResult)
                            : TryEvaluateCashFlowValue(function, args, out paymentResult);
                        if (!evaluated) return false;
                        if (paymentResult.IsError) { error = paymentResult; return false; }
                        result = paymentResult.Number!.Value;
                        return true;
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
                        if (!TryEvaluateMatchValue(function, args, out FormulaArgumentValue value)) return false;
                        if (value.IsError) { error = value; return false; }
                        if (!value.Number.HasValue) return false;
                        result = value.Number.Value;
                        return true;
                    }

                    if (IsSupportedTextFunction(function)) {
                        if (!TryEvaluateTextFunction(function, args, out FormulaArgumentValue textFunctionResult)) return false;
                        if (textFunctionResult.IsError) { error = textFunctionResult; return false; }
                        if (!textFunctionResult.Number.HasValue) return false;
                        result = textFunctionResult.Number.Value;
                        return true;
                    }

                    if (function == "AVERAGEA" || function == "MINA" || function == "MAXA") {
                        if (!TryEvaluateAValueAggregate(function, args, out FormulaArgumentValue aggregate) || !aggregate.Number.HasValue) return false;
                        result = aggregate.Number.Value;
                        return true;
                    }

                    if (function == "LARGE" || function == "SMALL") {
                        return TryEvaluateRankedAggregate(function, args, out result);
                    }

                    if (!TryResolveFormulaArguments(args, out var values)) {
                        return false;
                    }
                    foreach (FormulaArgumentValue value in values) {
                        if (value.IsUnresolvedFormula) return false;
                    }
                    if (_formulaReferenceError.IsError) { error = _formulaReferenceError; return false; }

                    if (function == "COUNTA") {
                        result = values.Count(v => v.HasValue || !string.IsNullOrEmpty(v.Text));
                        return true;
                    }

                    if (function != "COUNT") {
                        foreach (FormulaArgumentValue value in values) {
                            if (value.IsError) { error = value; return false; }
                        }
                    }


                    bool numericReferencesOnly = function is "SUM" or "AVERAGE" or "MIN" or "MAX" or "COUNT" or "PRODUCT" or "MEDIAN" or "SUMSQ";
                    var numbers = new List<double>(values.Count);
                    foreach (FormulaArgumentValue argument in values) {
                        FormulaArgumentValue v = numericReferencesOnly ? argument : NormalizeLookupResult(argument);
                        if (v.Number.HasValue && (!numericReferencesOnly || v.IsNumericAggregateValue)) numbers.Add(v.Number.Value);
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

                    if (function == "ROUND") {
                        if (numbers.Count != 2 || !TryGetSupportedDecimalPlaces(numbers[1], out int digits)) {
                            return false;
                        }

                        result = RoundAtDigits(numbers[0], digits, MidpointRounding.AwayFromZero);
                        return true;
                    }

                    if (function == "MROUND") {
                        if (numbers.Count != 2 || !TryEvaluateMRound(numbers[0], numbers[1], out result)) {
                            return false;
                        }

                        return true;
                    }

                    if (function == "CEILING.MATH" || function == "FLOOR.MATH") {
                        if (numbers.Count < 1 || numbers.Count > 3 || !TryEvaluateMathRoundFunction(function, numbers, out result)) {
                            return false;
                        }

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

                    if (numbers.Count == 0) {
                        if (function is "SUM" or "SUMSQ" or "MIN" or "MAX" or "PRODUCT") return true;
                        if (function == "AVERAGE") error = FormulaArgumentValue.Error("#DIV/0!");
                        else if (function == "MEDIAN") error = FormulaArgumentValue.Error("#NUM!");
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

                if (ExcelFormulaExpressionParser.TryParseArithmetic(formula, out ExcelFormulaBinaryExpressionSyntax? binary)) {
                    if (!TryResolveNumericOperand(binary!.Left, out double left)
                        || !TryResolveNumericOperand(binary.Right, out double right)) {
                        return false;
                    }

                    switch (binary.Operator) {
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

            return false;
        }

    }
}
