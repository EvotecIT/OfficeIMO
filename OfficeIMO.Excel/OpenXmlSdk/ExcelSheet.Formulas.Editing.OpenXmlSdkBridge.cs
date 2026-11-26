using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Marks all formula cells on this sheet dirty.
        /// </summary>
        public void InvalidateFormulas() {
            WriteLock(() => {
                foreach (var formula in ExcelWorksheetScan.CellFormulas(WorksheetRoot)) {
                    formula.CalculateCell = true;
                }
            });
        }

        /// <summary>
        /// Applies deferred defined-name formula marks. It runs from package access, which may already hold the
        /// document lock, so it marks the cells without taking it.
        /// </summary>
        /// <summary>
        /// Defers formula marks owed by defined-name edits around a literal value write. A sheet with tables, or one whose
        /// worksheet is not loaded, keeps the immediate marks: a header write can regenerate totals formulas.
        /// </summary>
        private ExcelDocument.DefinedNameFormulaMarkDeferral DeferDefinedNameFormulaMarksForValueWrite() {
            if (!_excelDocument.HasPendingDefinedNameFormulaMarks) return default;
            // Read through the unguarded field: the guarded one would apply the marks this defers.
            WorksheetPart part = BoundWorksheetPart;
            if (!part.IsRootElementLoaded || part.Worksheet is not Worksheet worksheet
                || worksheet.GetFirstChild<TableParts>() != null) return default;
            return _excelDocument.DeferDefinedNameFormulaMarks();
        }

        internal void MarkFormulasForRecalculation() {
            foreach (var formula in ExcelWorksheetScan.CellFormulas(WorksheetRoot)) {
                formula.CalculateCell = true;
            }
            MarkRequiresSavePreparation();
        }

        /// <summary>
        /// Removes cached values from formula cells on this sheet.
        /// </summary>
        public void ClearCachedFormulaResults() {
            WriteLock(() => {
                if (HasTableDefinitions && ExcelWorksheetScan.FormulaCells(WorksheetRoot).Any(FormulaTargetsTableHeader))
                    throw new InvalidOperationException("Clearing these formula caches would erase a table-column header.");
                bool changed = false;
                var arrays = WorksheetRoot.Descendants<CellFormula>()
                    .Where(formula => formula.FormulaType?.Value == CellFormulaValues.Array)
                    .Select(formula => formula.Reference?.Value).Where(reference => reference != null)
                    .Select(reference => ExcelA1.TryParseRange(reference!, out int top, out int left, out int bottom, out int right)
                        ? (Top: top, Left: left, Bottom: bottom, Right: right) : default).ToArray();
                var arrayIndex = new ExcelCellRangeIndex<bool>(arrays.Select(range => (range.Top, range.Left, range.Bottom, range.Right, true)));
                foreach (var cell in ExcelWorksheetScan.Cells(WorksheetRoot)) {
                    bool arrayMember = ExcelWorksheetScan.FindFormula(cell) == null && arrays.Length != 0
                        && cell.CellReference?.Value is string address
                        && TryParseCellReference(address, out int row, out int column)
                        && arrayIndex.At(row, column).Any();
                    if (ExcelWorksheetScan.FindFormula(cell) == null && !arrayMember) continue;
                    if (ExcelWorksheetScan.FindValue(cell) != null || cell.ValueMetaIndex != null) {
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

            });
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
                List<Cell> formulaCells = ExcelWorksheetScan.FormulaCells(WorksheetRoot).ToList();
                var dependencyInspectionContext = new FormulaDependencyInspectionContext(
                    this,
                    formulaCells.ConvertAll(cell => (ExcelFormulaCell)cell),
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
                using var context = new FormulaCalculationContext();
                var previousCalculationContext = _formulaCalculationContext;
                var previousCache = _formulaEvaluationCache;
                var previousDepthCache = _formulaEvaluationDepthCache;
                var previousStack = _formulaEvaluationStack;
                var previousDepthFrames = _formulaEvaluationDepthFrames;
                var previousGuardState = _formulaEvaluationGuardState;
                var previousSharedDefinitions = _formulaEvaluationSharedDefinitions;
                var previousSharedDefinitionsBySheet = _formulaEvaluationSharedDefinitionsBySheet;
                _formulaEvaluationCache = context.Cache;
                _formulaCalculationContext = context;
                _formulaEvaluationDepthCache = new Dictionary<long, int>();
                _formulaEvaluationStack = new HashSet<long>();
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
                            GetFormulaCachedText(cell),
                            packageRequestsRecalculation
                                || (ExcelWorksheetScan.FindFormula(cell)!.CalculateCell?.Value ?? false)
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
                    _formulaCalculationContext = previousCalculationContext;
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
            CellFormula? formula = ExcelWorksheetScan.FindFormula(cell);
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
        /// Sets a shared-free array formula over a range. The top-left cell owns the formula metadata.
        /// </summary>
        public void SetArrayFormula(string a1Range, string formula) => SetArrayFormulaCore(a1Range, formula, dynamic: false);

        /// <summary>
        /// Authors a dynamic array at one anchor cell. Calculation determines its bounded spill range.
        /// </summary>
        public void SetDynamicArrayFormula(string a1Cell, string formula) {
            var (row, column) = ExcelA1.ParseCellRef(a1Cell);
            if (row < 1 || row > ExcelA1.MaxRows || column < 1 || column > ExcelA1.MaxColumns)
                throw new ArgumentOutOfRangeException(nameof(a1Cell));
            string anchor = ExcelA1.CellReference(row, column);
            SetArrayFormulaCore(anchor + ":" + anchor, formula, dynamic: true);
        }

        private void SetArrayFormulaCore(string a1Range, string formula, bool dynamic) {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentNullException(nameof(formula));
            string safeFormula = Utilities.ExcelSanitizer.SanitizeFormula(formula, qualifyArrayFunctions: true);
            var (r1, c1, r2, c2) = ExcelA1.ParseRange(a1Range);
            WriteLock(() => {
                EnsureTableHeaderFormulaRangeWritable(r1, c1, r2, c2);
                foreach (var cell in ExcelWorksheetScan.FormulaCells(WorksheetRoot).Where(c => ExcelWorksheetScan.FindFormula(c)?.FormulaType?.Value == CellFormulaValues.Array).ToList()) {
                    string? reference = ExcelWorksheetScan.FindFormula(cell)?.Reference?.Value;
                    if (!string.IsNullOrWhiteSpace(reference)
                        && TryFixedArrayBounds(reference!, out int existingR1, out int existingC1, out int existingR2, out int existingC2)
                        && RangesOverlapInclusive((r1, c1, r2, c2), (existingR1, existingC1, existingR2, existingC2))) {
                        throw new InvalidOperationException($"Array formula range '{a1Range}' overlaps existing array formula range '{reference}'.");
                    }
                }

                var topLeft = GetCell(r1, c1);
                bool retainsCachedValue = ExcelWorksheetScan.FindValue(topLeft) != null;
                ClearCellValueMetadata(topLeft);
                topLeft.CellMetaIndex = dynamic ? EnsureDynamicArrayMetadata() : null;
                topLeft.CellFormula = new CellFormula(safeFormula) {
                    FormulaType = CellFormulaValues.Array,
                    Reference = a1Range
                };
                if (retainsCachedValue) {
                    ExcelWorksheetScan.FindFormula(topLeft)!.CalculateCell = true;
                }
                _excelDocument.MarkFormulaAuthored(
                    _worksheetPart,
                    ExcelA1.CellReference(r1, c1),
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
            });
        }

        internal void SetLegacyArrayFormula(string a1Range, string formula, System.Threading.CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentNullException(nameof(formula));
            string safeFormula = Utilities.ExcelSanitizer.SanitizeFormula(formula);
            int r2;
            int c2;
            if (!ExcelA1.TryParseRange(a1Range, out int r1, out int c1, out r2, out c2)) {
                (r1, c1) = ExcelA1.ParseCellRef(a1Range);
                if (r1 <= 0 || c1 <= 0) {
                    throw new ArgumentException($"Invalid A1 range '{a1Range}'.", nameof(a1Range));
                }
                r2 = r1;
                c2 = c1;
            }

            WriteLock(() => {
                for (int row = r1; row <= r2; row++) {
                    for (int column = c1; column <= c2; column++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        var member = GetCell(row, column);
                        ClearCellValueMetadata(member);
                        if (row != r1 || column != c1) member.CellFormula = null;
                    }
                }
                var topLeft = GetCell(r1, c1);
                bool retainsCachedValue = ExcelWorksheetScan.FindValue(topLeft) != null;
                topLeft.CellFormula = new CellFormula(safeFormula) {
                    FormulaType = CellFormulaValues.Array,
                    Reference = a1Range
                };
                if (retainsCachedValue) {
                    ExcelWorksheetScan.FindFormula(topLeft)!.CalculateCell = true;
                }
                _excelDocument.MarkFormulaAuthored(
                    _worksheetPart,
                    ExcelA1.CellReference(r1, c1),
                    retainedCachedValue: retainsCachedValue);
            });
        }

        /// <summary>
        /// Clears an array formula whose reference overlaps the supplied range or cell.
        /// </summary>
        public void ClearArrayFormula(string a1RangeOrCell) {
            var bounds = a1RangeOrCell.IndexOf(':') >= 0
                ? ExcelA1.ParseRange(a1RangeOrCell)
                : CellAsRange(a1RangeOrCell);

            WriteLock(() => {
                var arrayCells = ExcelWorksheetScan.FormulaCells(WorksheetRoot)
                    .Where(c => ExcelWorksheetScan.FindFormula(c)?.FormulaType?.Value == CellFormulaValues.Array).ToList();
                foreach (Cell array in arrayCells) {
                    if (TryFixedArrayBounds(ExcelWorksheetScan.FindFormula(array)?.Reference?.Value ?? string.Empty,
                        out int top, out int left, out int bottom, out int right)
                        && RangesOverlapInclusive(bounds, (top, left, bottom, right)))
                        EnsureTableHeaderValuesNotCleared(top, left, bottom, right);
                }
                foreach (var cell in arrayCells) {
                    string? reference = ExcelWorksheetScan.FindFormula(cell)?.Reference?.Value;
                    if (!string.IsNullOrWhiteSpace(reference)
                        && TryFixedArrayBounds(reference!, out int existingR1, out int existingC1, out int existingR2, out int existingC2)
                        && RangesOverlapInclusive(bounds, (existingR1, existingC1, existingR2, existingC2))) {
                        foreach (Cell spillCell in ExcelWorksheetScan.Cells(WorksheetRoot).ToList()) {
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
            });
        }
    }
}
