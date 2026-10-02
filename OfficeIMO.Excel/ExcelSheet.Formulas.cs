using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
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
        private IReadOnlyDictionary<uint, SharedFormulaDefinition>? _formulaEvaluationSharedDefinitions;
        private Dictionary<string, IReadOnlyDictionary<uint, SharedFormulaDefinition>>? _formulaEvaluationSharedDefinitionsBySheet;
        private string? _formulaEvaluationCellReference;

        internal sealed class FormulaEvaluationGuardState {
            internal bool DependencyGuardBlocked { get; set; }
        }

        internal sealed class FormulaEvaluationDepthFrame {
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
                    if (cell.CellValue != null || cell.ValueMetaIndex != null) {
                        cell.CellValue = null;
                        ClearCellValueMetadataAttribute(cell);
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
        public int RecalculateSupportedFormulas() => RecalculateSupportedFormulas(new FormulaCalculationContext());

        internal int RecalculateSupportedFormulas(FormulaCalculationContext context) {
            MaterializePendingDirectCellValues();

            int count = 0;
            WriteLock(() => {
                MaterializePendingDirectCellValues();
                long formulaInputMutationVersion = _excelDocument.CaptureFormulaInputMutationVersion();

                var previousCache = _formulaEvaluationCache;
                var previousDepthCache = _formulaEvaluationDepthCache;
                var previousStack = _formulaEvaluationStack;
                var previousDepthFrames = _formulaEvaluationDepthFrames;
                var previousGuardState = _formulaEvaluationGuardState;
                var previousSharedDefinitions = _formulaEvaluationSharedDefinitions;
                var previousSharedDefinitionsBySheet = _formulaEvaluationSharedDefinitionsBySheet;
                _formulaEvaluationCache = context.Cache;
                _formulaEvaluationDepthCache = context.DepthCache;
                _formulaEvaluationStack = context.Stack;
                _formulaEvaluationDepthFrames = context.DepthFrames;
                _formulaEvaluationGuardState = context.GuardState;
                _formulaEvaluationSharedDefinitions = BuildSharedFormulaDefinitions();
                _formulaEvaluationSharedDefinitionsBySheet = new Dictionary<string, IReadOnlyDictionary<uint, SharedFormulaDefinition>>(
                    StringComparer.OrdinalIgnoreCase) {
                    [Name] = _formulaEvaluationSharedDefinitions
                };
                bool changed = false;
                bool allFormulasEvaluated = true;

                try {
                    if (_excelDocument.WorkbookPartRoot.CellMetadataPart != null) PlanDynamicArrayOwners();
                    foreach (var cell in WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula != null).ToList()) {
                        _formulaEvaluationGuardState.DependencyGuardBlocked = false;
                        if (!TryEvaluateFormulaCellValue(cell, out FormulaArgumentValue result)) {
                            allFormulasEvaluated = false;
                            if (_formulaEvaluationGuardState.DependencyGuardBlocked) {
                                if (cell.CellValue != null || cell.ValueMetaIndex != null) {
                                    cell.CellValue = null;
                                    ClearCellValueMetadataAttribute(cell);
                                    changed = true;
                                }

                                if (cell.CellFormula!.CalculateCell?.Value != true) {
                                    cell.CellFormula.CalculateCell = true;
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
                        cell.CellFormula!.CalculateCell = false;
                        _excelDocument.MarkFormulaCellRecalculated(
                            _worksheetPart,
                            cell.CellReference?.Value ?? string.Empty,
                            formulaInputMutationVersion);
                        changed = true;
                        count++;
                    }
                } finally {
                    _formulaEvaluationCache = previousCache;
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
                    _excelDocument.MarkFormulaSheetRecalculated(_worksheetPart, formulaInputMutationVersion);
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

        private bool TryEvaluateFormulaCellValue(
            Cell cell,
            out FormulaArgumentValue result,
            IReadOnlyDictionary<uint, SharedFormulaDefinition>? sharedFormulaDefinitions = null) {
            result = default;
            if (cell.CellFormula == null) {
                return false;
            }

            string formula = ResolveCellFormulaText(
                cell,
                sharedFormulaDefinitions ?? _formulaEvaluationSharedDefinitions);

            string? reference = NormalizeFormulaCellReference(cell.CellReference?.Value);
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
                    return cell.CellFormula.FormulaType?.Value == CellFormulaValues.Array
                        ? TryEvaluateAuthoredFormulaCell(cell, formula, out result)
                        : TryEvaluateFormulaValue(formula, out result);
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

                var depthFrame = new FormulaEvaluationDepthFrame();
                _formulaEvaluationDepthFrames.Push(depthFrame);
                bool evaluated = false;
                int evaluationDepth = 0;
                try {
                    if (!(cell.CellFormula.FormulaType?.Value == CellFormulaValues.Array
                        ? TryEvaluateAuthoredFormulaCell(cell, formula, out result)
                        : TryEvaluateFormulaValue(formula, out result))) {
                        return false;
                    }
                    if (depthFrame.DependencyGuardBlocked) {
                        return false;
                    }

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

        private void SetFormulaCachedValue(Cell cell, FormulaArgumentValue result) {
            cell.ValueMetaIndex = null;
            if (result.IsError && TryWriteRichFormulaError(cell, result.ErrorCode)) return;
            if (result.IsBoolean) {
                cell.CellValue = new CellValue(result.Number == 0 ? "0" : "1");
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean;
                return;
            }
            if (!result.IsError && result.SourceCellKind == ExcelCellDataKind.Text && result.Text != null) {
                cell.CellValue = new CellValue(result.Text);
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                return;
            }
            if (result.Number.HasValue) {
                cell.CellValue = new CellValue(InvariantNumberText.Get(result.Number.Value));
                cell.DataType = result.IsBoolean ? DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean
                    : DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
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
            return _excelDocument.ExecuteReadAfterMaterializing(() => {
                var formulas = new List<ExcelFormulaCellInfo>();
                FormulaDependencyAliasCatalog dependencyAliases = GetFormulaDependencyAliases();
                FormulaDependencyTableCatalog dependencyTables = GetFormulaDependencyTables();
                IReadOnlyDictionary<uint, SharedFormulaDefinition> sharedFormulaDefinitions = BuildSharedFormulaDefinitions();
                List<Cell> formulaCells = WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula != null).ToList();
                var dependencyInspectionContext = new FormulaDependencyInspectionContext(
                    this,
                    formulaCells,
                    sharedFormulaDefinitions);
                CellMetadataPart? metadataPart = _excelDocument.WorkbookPartRoot.CellMetadataPart;
                if (metadataPart != null && !metadataPart.IsRootElementLoaded) {
                    ValidateInCellImageMetadataPart(metadataPart, "Cell metadata");
                }
                Metadata? metadata = metadataPart?.Metadata;
                CalculationProperties? calculation = WorkbookRoot.GetFirstChild<CalculationProperties>();
                bool packageRequestsRecalculation = calculation?.FullCalculationOnLoad?.Value == true
                    || calculation?.ForceFullCalculation?.Value == true
                    || WorksheetRoot.GetFirstChild<SheetCalculationProperties>()?.FullCalculationOnLoad?.Value == true;
                var previousCache = _formulaEvaluationCache;
                var previousDepthCache = _formulaEvaluationDepthCache;
                var previousStack = _formulaEvaluationStack;
                var previousDepthFrames = _formulaEvaluationDepthFrames;
                var previousGuardState = _formulaEvaluationGuardState;
                var previousSharedDefinitions = _formulaEvaluationSharedDefinitions;
                var previousSharedDefinitionsBySheet = _formulaEvaluationSharedDefinitionsBySheet;
                _formulaEvaluationCache = new Dictionary<string, FormulaArgumentValue>(StringComparer.OrdinalIgnoreCase);
                _formulaEvaluationDepthCache = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                _formulaEvaluationStack = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _formulaEvaluationDepthFrames = new Stack<FormulaEvaluationDepthFrame>();
                _formulaEvaluationGuardState = new FormulaEvaluationGuardState();
                _formulaEvaluationSharedDefinitions = sharedFormulaDefinitions;
                _formulaEvaluationSharedDefinitionsBySheet = new Dictionary<string, IReadOnlyDictionary<uint, SharedFormulaDefinition>>(
                    StringComparer.OrdinalIgnoreCase) {
                    [Name] = sharedFormulaDefinitions
                };
                try {
                    if (metadataPart != null) PlanDynamicArrayOwners();
                    foreach (Cell cell in formulaCells) {
                        _formulaEvaluationGuardState.DependencyGuardBlocked = false;
                        string formula = ResolveCellFormulaText(cell, sharedFormulaDefinitions);
                        string cellReference = cell.CellReference?.Value ?? string.Empty;
                        bool supported = TryEvaluateFormulaCellValue(cell, out _, sharedFormulaDefinitions);
                        IReadOnlyList<string> dependencies = GetFormulaDependencies(
                            cell.CellReference?.Value,
                            formula,
                            dependencyAliases,
                            dependencyTables);
                        IReadOnlyList<string> dependencyIssues = GetFormulaDependencyIssues(
                            cell.CellReference?.Value,
                            dependencies,
                            dependencyInspectionContext);
                        ExcelFormulaArrayInfo? arrayInfo = CreateFormulaArrayInfo(cell, metadata);
                        formulas.Add(new ExcelFormulaCellInfo(
                            Name,
                            cellReference,
                            formula,
                            cell.CellValue?.Text,
                            packageRequestsRecalculation
                                || (cell.CellFormula!.CalculateCell?.Value ?? false)
                                || _excelDocument.HasFormulaInputMutationsAfterFormulaBaseline(_worksheetPart, cellReference)
                                || HasFormulaDependencyMutationsAfterBaseline(cellReference, dependencies),
                            supported,
                            supported ? null : GetUnsupportedFormulaReason(formula),
                            dependencies,
                            dependencyIssues,
                            arrayInfo));
                    }
                } finally {
                    _formulaEvaluationCache = previousCache;
                    _formulaEvaluationDepthCache = previousDepthCache;
                    _formulaEvaluationStack = previousStack;
                    _formulaEvaluationDepthFrames = previousDepthFrames;
                    _formulaEvaluationGuardState = previousGuardState;
                    _formulaEvaluationSharedDefinitions = previousSharedDefinitions;
                    _formulaEvaluationSharedDefinitionsBySheet = previousSharedDefinitionsBySheet;
                }

                return formulas;
            });
        }

        private bool HasFormulaDependencyMutationsAfterBaseline(
            string cellReference,
            IReadOnlyList<string> dependencies) {
            if (dependencies.Count == 0) return false;
            long baseline = _excelDocument.GetFormulaDependencyBaseline(_worksheetPart, cellReference);
            foreach (string dependency in dependencies) {
                if (!TryResolveFormulaDependencyReference(
                    dependency,
                    out ExcelSheet dependencySheet,
                    out int firstRow,
                    out int firstColumn,
                    out int lastRow,
                    out int lastColumn,
                    out _)) {
                    continue;
                }
                if (_excelDocument.HasFormulaDependencyMutationAfter(
                    dependencySheet.WorksheetPart,
                    firstRow,
                    firstColumn,
                    lastRow,
                    lastColumn,
                    baseline)) {
                    return true;
                }
            }
            return false;
        }

        private static ExcelFormulaArrayInfo? CreateFormulaArrayInfo(
            Cell cell,
            Metadata? metadata) {
            CellFormula? formula = cell.CellFormula;
            if (formula?.FormulaType?.Value != CellFormulaValues.Array
                || string.IsNullOrWhiteSpace(formula.Reference?.Value)) {
                return null;
            }

            uint? metadataIndex = null;
            foreach (var attribute in cell.GetAttributes()) {
                if (string.Equals(attribute.LocalName, "cm", StringComparison.OrdinalIgnoreCase)
                    && uint.TryParse(attribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed)) {
                    metadataIndex = parsed;
                    break;
                }
            }

            bool dynamic = false;
            bool collapsed = false;
            if (metadataIndex is uint oneBasedMetadataIndex && oneBasedMetadataIndex > 0) {
                MetadataBlock? cellBlock = metadata?.GetFirstChild<CellMetadata>()?
                    .Elements<MetadataBlock>()
                    .ElementAtOrDefault(oneBasedMetadataIndex > int.MaxValue
                        ? -1
                        : (int)oneBasedMetadataIndex - 1);
                MetadataType[] types = metadata?.GetFirstChild<MetadataTypes>()?
                    .Elements<MetadataType>()
                    .ToArray() ?? Array.Empty<MetadataType>();
                foreach (MetadataRecord record in cellBlock?.Elements<MetadataRecord>() ?? Enumerable.Empty<MetadataRecord>()) {
                    uint oneBasedTypeIndex = record.TypeIndex?.Value ?? 0U;
                    if (oneBasedTypeIndex == 0 || oneBasedTypeIndex > types.Length) continue;
                    string? typeName = types[oneBasedTypeIndex - 1].Name?.Value;
                    if (!string.Equals(typeName, "XLDAPR", StringComparison.OrdinalIgnoreCase)) continue;
                    uint valueIndex = record.Val?.Value ?? uint.MaxValue;
                    FutureMetadata? future = metadata?.Elements<FutureMetadata>()
                        .FirstOrDefault(item => string.Equals(item.Name?.Value, typeName, StringComparison.OrdinalIgnoreCase));
                    FutureMetadataBlock? futureBlock = future?.Elements<FutureMetadataBlock>()
                        .ElementAtOrDefault(valueIndex > int.MaxValue ? -1 : (int)valueIndex);
                    OpenXmlElement? properties = futureBlock?.Descendants()
                        .FirstOrDefault(item => string.Equals(item.LocalName, "dynamicArrayProperties", StringComparison.OrdinalIgnoreCase));
                    if (properties == null) continue;
                    string? dynamicValue = properties.GetAttributes()
                        .FirstOrDefault(item => string.Equals(item.LocalName, "fDynamic", StringComparison.OrdinalIgnoreCase)).Value;
                    dynamic = !string.Equals(dynamicValue, "0", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(dynamicValue, "false", StringComparison.OrdinalIgnoreCase);
                    string? collapsedValue = properties.GetAttributes()
                        .FirstOrDefault(item => string.Equals(item.LocalName, "fCollapsed", StringComparison.OrdinalIgnoreCase)).Value;
                    collapsed = dynamic && (string.Equals(collapsedValue, "1", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(collapsedValue, "true", StringComparison.OrdinalIgnoreCase));
                    break;
                }
            }
            return new ExcelFormulaArrayInfo(formula.Reference!.Value!, dynamic, collapsed, metadataIndex);
        }

        /// <summary>
        /// Returns the formula text from a cell, if present.
        /// </summary>
        public string? GetFormulaText(int row, int column) {
            Cell? cell = TryGetExistingCell(row, column);
            return cell?.CellFormula == null ? null : ResolveCellFormulaText(cell);
        }

        /// <summary>
        /// Tries to return a formula cell's cached value.
        /// </summary>
        public bool TryGetCachedFormulaValue(int row, int column, out string? value) {
            var cell = TryGetExistingCell(row, column);
            value = cell?.CellFormula == null ? null : ResolveRichValueError(cell, cell.CellValue?.Text);
            return value != null;
        }

        /// <summary>
        /// Sets a shared-free array formula over a range. The top-left cell owns the formula metadata.
        /// </summary>
        public void SetArrayFormula(string a1Range, string formula) => SetArrayFormulaCore(a1Range, formula, dynamic: false);

        /// <summary>
        /// Authors a dynamic array at one anchor cell. Calculation determines its bounded spill range.
        /// </summary>
        public void SetDynamicArrayFormula(string a1Cell, string formula) {
            var (row, column) = A1.ParseCellRef(a1Cell);
            if (row < 1 || row > A1.MaxRows || column < 1 || column > A1.MaxColumns)
                throw new ArgumentOutOfRangeException(nameof(a1Cell));
            string anchor = A1.CellReference(row, column);
            SetArrayFormulaCore(anchor + ":" + anchor, formula, dynamic: true);
        }

        private void SetArrayFormulaCore(string a1Range, string formula, bool dynamic) {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentNullException(nameof(formula));
            string safeFormula = Utilities.ExcelSanitizer.SanitizeFormula(formula);
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            WriteLock(() => {
                foreach (var cell in WorksheetRoot.Descendants<Cell>().Where(c => c.CellFormula?.FormulaType?.Value == CellFormulaValues.Array).ToList()) {
                    string? reference = cell.CellFormula?.Reference?.Value;
                    if (!string.IsNullOrWhiteSpace(reference)
                        && TryFixedArrayBounds(reference!, out int existingR1, out int existingC1, out int existingR2, out int existingC2)
                        && RangesOverlapInclusive((r1, c1, r2, c2), (existingR1, existingC1, existingR2, existingC2))) {
                        throw new InvalidOperationException($"Array formula range '{a1Range}' overlaps existing array formula range '{reference}'.");
                    }
                }

                var topLeft = GetCell(r1, c1);
                bool retainsCachedValue = topLeft.CellValue != null;
                ClearCellValueMetadata(topLeft);
                topLeft.CellMetaIndex = dynamic ? EnsureDynamicArrayMetadata() : null;
                topLeft.CellFormula = new CellFormula(QualifyAuthoredArrayFunctions(safeFormula)) {
                    FormulaType = CellFormulaValues.Array,
                    Reference = a1Range
                };
                if (retainsCachedValue) {
                    topLeft.CellFormula.CalculateCell = true;
                }
                _excelDocument.MarkFormulaAuthored(
                    _worksheetPart,
                    A1.CellReference(r1, c1),
                    retainedCachedValue: retainsCachedValue);
                for (int row = r1; row <= r2; row++) {
                    for (int column = c1; column <= c2; column++) {
                        if (row == r1 && column == c1) continue;
                        var cell = GetCell(row, column);
                        ClearCellValueMetadata(cell);
                        cell.CellFormula = null;
                        cell.CellValue = null;
                    }
                }

                InvalidateDynamicArrayWriteIndex();
                WorksheetRoot.Save();
            });
        }

        internal void SetLegacyArrayFormula(string a1Range, string formula) {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentNullException(nameof(formula));
            string safeFormula = Utilities.ExcelSanitizer.SanitizeFormula(formula);
            int r2;
            int c2;
            if (!A1.TryParseRange(a1Range, out int r1, out int c1, out r2, out c2)) {
                (r1, c1) = A1.ParseCellRef(a1Range);
                if (r1 <= 0 || c1 <= 0) {
                    throw new ArgumentException($"Invalid A1 range '{a1Range}'.", nameof(a1Range));
                }
                r2 = r1;
                c2 = c1;
            }

            WriteLock(() => {
                for (int row = r1; row <= r2; row++) {
                    for (int column = c1; column <= c2; column++) {
                        ClearCellValueMetadata(GetCell(row, column));
                    }
                }
                var topLeft = GetCell(r1, c1);
                bool retainsCachedValue = topLeft.CellValue != null;
                topLeft.CellFormula = new CellFormula(safeFormula) {
                    FormulaType = CellFormulaValues.Array,
                    Reference = a1Range
                };
                if (retainsCachedValue) {
                    topLeft.CellFormula.CalculateCell = true;
                }
                _excelDocument.MarkFormulaAuthored(
                    _worksheetPart,
                    A1.CellReference(r1, c1),
                    retainedCachedValue: retainsCachedValue);
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
                        && TryFixedArrayBounds(reference!, out int existingR1, out int existingC1, out int existingR2, out int existingC2)
                        && RangesOverlapInclusive(bounds, (existingR1, existingC1, existingR2, existingC2))) {
                        foreach (Cell spillCell in WorksheetRoot.Descendants<Cell>().ToList()) {
                            if (!TryParseCellReference(spillCell.CellReference?.Value ?? "", out int row, out int column)
                                || row < existingR1 || row > existingR2 || column < existingC1 || column > existingC2)
                                continue;
                            spillCell.CellFormula = null;
                            spillCell.CellValue = null;
                            ClearCellValueMetadataAttribute(spillCell);
                            spillCell.DataType = null;
                            if (ReferenceEquals(spillCell, cell)) spillCell.CellMetaIndex = null;
                            DynamicSpillCacheSnapshots.Remove(DynamicCellKey(row, column));
                        }
                        SpillOwnership.WrittenOwners.Remove(DynamicCellKey(existingR1, existingC1));
                    }
                }

                InvalidateDynamicArrayWriteIndex();
                WorksheetRoot.Save();
            });
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
                    string function = functionCall.Name.ToUpperInvariant();
                    string args = functionCall.Arguments;
                    if (function == "MATCH" || function == "XMATCH") return TryEvaluateMatchValue(function, args, out result);
                    if (function is "INT" or "MOD" or "SQRT" or "SIGN" or "TRUNC" or "ROUNDUP" or "ROUNDDOWN" or "EXP" or "LN" or "LOG" or "LOG10" or "CEILING" or "FLOOR" or "EVEN" or "ODD" or "FACT" or "COMBIN") return TryEvaluateScalarMathValue(function, args, out result);
                    if (function == "PROB") return TryEvaluateProbabilityValue(args, out result);
                    if (function == "RANDBETWEEN") return TryEvaluateRandomBetweenValue(args, out result);
                    if (function == "OFFSET") return TryEvaluateOffsetValue(args, out result);
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
                    if (function == "DATE" && TryEvaluateDateValue(args, out result)) return true;
                    if (function == "DATEDIF" && TryEvaluateDateDifValue(args, out result)) return true;
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

                    if ((function == "ISBLANK" || function == "ISNUMBER" || function == "ISLOGICAL" || function == "ISTEXT" || function == "ISERROR" || function == "ISERR" || function == "ISNA" || function == "ISFORMULA")
                        && TryEvaluateInfoFunction(function, args, out result)) {
                        return true;
                    }

                    if ((function == "AVERAGEA" || function == "MINA" || function == "MAXA")
                        && TryEvaluateAValueAggregate(function, args, out result)) {
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

            if (functionCall == null && TryEvaluateCustomFormulaFunction(formula, out result)) {
                return true;
            }

            if (TryEvaluateSingleReferenceFormulaValue(formula, out result)) {
                return true;
            }

            if (TryEvaluateFormulaCore(formula, out double numeric, out FormulaArgumentValue error)) {
                bool isBoolean = functionCall != null && (functionCall.Name.Equals("AND", StringComparison.OrdinalIgnoreCase)
                    || functionCall.Name.Equals("OR", StringComparison.OrdinalIgnoreCase) || functionCall.Name.Equals("NOT", StringComparison.OrdinalIgnoreCase));
                result = new FormulaArgumentValue(numeric, InvariantNumberText.Get(numeric), isBoolean: isBoolean);
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
                    string function = functionCall.Name.ToUpperInvariant();
                    string args = functionCall.Arguments;
                    if (function is "INT" or "MOD" or "SQRT" or "SIGN" or "TRUNC" or "ROUNDUP" or "ROUNDDOWN" or "EXP" or "LN" or "LOG" or "LOG10" or "CEILING" or "FLOOR" or "EVEN" or "ODD" or "FACT" or "COMBIN") {
                        if (!TryEvaluateScalarMathValue(function, args, out FormulaArgumentValue value)) return false;
                        if (value.IsError) { error = value; return false; }
                        result = value.Number!.Value;
                        return true;
                    }
                    if (function == "PROB" || function == "RANDBETWEEN" || function == "OFFSET") {
                        bool evaluated = function == "PROB" ? TryEvaluateProbabilityValue(args, out FormulaArgumentValue value)
                            : function == "RANDBETWEEN" ? TryEvaluateRandomBetweenValue(args, out value)
                            : TryEvaluateOffsetValue(args, out value);
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
                        if (!TryEvaluateMatchValue(function, args, out FormulaArgumentValue value)) return false;
                        if (value.IsError) { error = value; return false; }
                        if (!value.Number.HasValue) return false;
                        result = value.Number.Value;
                        return true;
                    }

                    if (TryEvaluateTextFunction(function, args, out FormulaArgumentValue textFunctionResult)
                        && textFunctionResult.Number.HasValue) {
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

                    if (!TryResolveFormulaArguments(args, out var values) || values.Any(value => value.IsUnresolvedFormula)) {
                        return false;
                    }

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
                    var numbers = values.Where(v => v.Number.HasValue && (!numericReferencesOnly || v.IsNumericAggregateValue))
                        .Select(v => v.Number!.Value).ToList();
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
