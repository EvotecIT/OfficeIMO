namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>The sheet as formulas read and write it.</summary>
        internal ExcelFormulaSheetModel FormulaModel => _excelDocument.GetFormulaSheetModel(this);

        /// <summary>The sheet's identity for formula caches: the same for every wrapper of one worksheet.</summary>
        internal Uri FormulaSheetKey => FormulaModel.PartUri;

        /// <summary>Whether the cell at the position holds a formula of its own.</summary>
        internal bool HasFormulaCell(int row, int column) => FormulaModel.Find(row, column)?.HasFormula == true;

        /// <summary>The cell at the position as the formula engine reads it, or an empty handle.</summary>
        internal ExcelFormulaCell GetFormulaCell(int row, int column) => new(FormulaModel.Find(row, column));

        /// <summary>The cells of the sheet that hold a formula element, in document order.</summary>
        internal IEnumerable<ExcelFormulaCell> EnumerateFormulaCells() {
            foreach (ExcelFormulaCellRecord cell in FormulaModel.FormulaCells()) yield return new ExcelFormulaCell(cell);
        }

        /// <summary>Every stored cell of the sheet, in document order.</summary>
        internal IEnumerable<ExcelFormulaCell> EnumerateStoredFormulaCells() {
            foreach (ExcelFormulaCellRecord cell in FormulaModel.Cells()) yield return new ExcelFormulaCell(cell);
        }

        /// <summary>The references of the sheet's merged ranges.</summary>
        private IEnumerable<string> EnumerateMergedRanges() => FormulaModel.MergedRanges();

        /// <summary>The ranges of the sheet's tables.</summary>
        private IEnumerable<string> EnumerateTableRanges() {
            foreach (ExcelFormulaTable table in FormulaModel.Tables()) yield return table.Reference ?? string.Empty;
        }

        /// <summary>The effective coordinates of every cell of the sheet, including cells that omit their reference.</summary>
        internal IReadOnlyDictionary<ExcelFormulaCell, (int Row, int Column)> BuildEffectiveCellCoordinates() =>
            FormulaModel.EffectiveCoordinates();

        /// <summary>The cell's value as cell reads give it, from the model formulas read.</summary>
        private ExcelCellData GetFormulaValueSnapshot(int row, int column) => FormulaModel.ReadValue(row, column);

        /// <summary>Whether an array formula's cell metadata marks it as a dynamic array; reads the metadata once per walk.</summary>
        private bool IsDynamicArrayFormula(ExcelFormulaCell candidate, ref object? metadataState) {
            metadataState ??= (object?)_excelDocument.FormulaWorkbookModel.CellMetadata ?? DBNull.Value;
            return ExcelCellMetadataIndex.ArrayInfo(candidate, metadataState as ExcelCellMetadataIndex)?.IsDynamic == true;
        }

        /// <summary>Removes a formula's cached value and value metadata; true when there was one to remove.</summary>
        private static bool ClearFormulaCachedValue(ExcelFormulaCell handle) {
            ExcelFormulaCellRecord cell = handle.Record!;
            if (cell.ValueText == null && cell.ValueMetaIndex == null) return false;
            cell.SetValue(null, cell.ValueType);
            cell.ClearValueMetaIndex();
            return true;
        }

        private void MarkFormulaCellRecalculated(string reference, long mutationVersion) =>
            _excelDocument.MarkFormulaCellRecalculated(FormulaSheetKey, reference, mutationVersion);

        private void MarkFormulaSheetRecalculated(long mutationVersion) =>
            _excelDocument.MarkFormulaSheetRecalculated(FormulaSheetKey, mutationVersion);

        private void SetFormulaCachedValue(ExcelFormulaCell handle, FormulaArgumentValue result) {
            ExcelFormulaCellRecord cell = handle.Record!;
            // Each write leaves an equal value and type as they are.
            if (cell.ValueMetaIndex != null) cell.ClearValueMetaIndex();
            if (result.IsError && TryWriteRichFormulaError(handle, result.ErrorCode)) return;
            if (result.IsBoolean) {
                SetFormulaCachedText(cell, result.Number == 0 ? "0" : "1", ExcelCellValueType.Boolean);
                return;
            }
            if (!result.IsError && result.SourceCellKind == ExcelCellDataKind.Text && result.Text != null) {
                SetFormulaCachedText(cell, result.Text, ExcelCellValueType.String);
                return;
            }
            if (result.Number.HasValue) {
                SetFormulaCachedText(cell, InvariantNumberText.Get(result.Number.Value), ExcelCellValueType.Number);
                return;
            }

            if (result.IsError) {
                SetFormulaCachedText(cell, result.ErrorCode ?? "#VALUE!", ExcelCellValueType.Error);
                return;
            }

            if (result.Text != null) {
                SetFormulaCachedText(cell, result.Text, ExcelCellValueType.String);
            }
        }

        private static void SetFormulaCachedText(ExcelFormulaCellRecord cell, string text, ExcelCellValueType type) =>
            cell.SetValue(type == ExcelCellValueType.String ? Utilities.ExcelSanitizer.EncodeXString(text) : text, type);

        private void WriteFixedArrayFormulaCache(ExcelFormulaCell ownerHandle, FormulaArrayValue array) {
            ExcelFormulaCellRecord owner = ownerHandle.Record!;
            var (row, column) = ExcelA1.ParseCellRef(owner.Reference!);
            ExcelFormulaSheetModel model = FormulaModel;
            for (int r = 0; r < array.Rows; r++) {
                for (int c = 0; c < array.Columns; c++) {
                    ExcelFormulaCellRecord cell = r == 0 && c == 0 ? owner : model.GetOrCreate(row + r, column + c);
                    cell.ClearInlineString();
                    cell.ClearValueMetaIndex();
                    SetFormulaCachedValue(new ExcelFormulaCell(cell), array.Values[r * array.Columns + c]);
                }
            }
        }

        /// <summary>Whether the formula writes a table's literal header: its own cell, or an array range over one.</summary>
        private bool FormulaTargetsTableHeader(ExcelFormulaCell cell) {
            ExcelFormulaSheetModel model = FormulaModel;
            if (!model.HasTableHeaders) return false;
            if (cell.Record!.TryGetPosition(out int row, out int column) && model.IsTableHeader(row, column, row, column)) return true;
            return cell.FormulaKind == ExcelFormulaKind.Array
                && TryFixedArrayBounds(cell.FormulaReference ?? string.Empty, out int top, out int left, out int bottom, out int right)
                && model.IsTableHeader(top, left, bottom, right);
        }

        private static string? ResolveRichValueError(ExcelFormulaCell cell, string? fallback) => cell.Record!.ResolveRichValueError(fallback);

        // The dominant operands: numeric and shared-text cells need no typed snapshot, boxed value or error-literal check.
        private bool TryResolvePlainNumberCell(int row, int column, out FormulaArgumentValue value) {
            value = default;
            if (!FormulaModel.TryReadPlainValue(row, column, out double? number, out string? text)) return false;
            value = number.HasValue
                ? new FormulaArgumentValue(number.Value, text, sourceCellKind: ExcelCellDataKind.Number)
                : new FormulaArgumentValue(null, text ?? string.Empty, sourceCellKind: ExcelCellDataKind.Text);
            return true;
        }

        /// <summary>The workbook's sheet entries in order, each marked as a worksheet or not.</summary>
        internal List<ExcelFormulaSheetEntry> ReadFormulaSheetEntries() => _excelDocument.FormulaWorkbookModel.SheetEntries();

        /// <summary>The workbook's sheet names in order.</summary>
        internal List<string?> ReadFormulaSheetNames() => ReadFormulaSheetEntries().ConvertAll(entry => entry.Name);

        /// <summary>The workbook's defined names in order.</summary>
        internal List<ExcelFormulaDefinedName> ReadFormulaDefinedNames() => _excelDocument.FormulaWorkbookModel.DefinedNames();

        /// <summary>A sheet wrapper for a workbook sheet entry.</summary>
        private ExcelSheet CreateFormulaSheetWrapper(ExcelFormulaSheetEntry entry) => _excelDocument.FormulaWorkbookModel.CreateSheet(entry);

        /// <summary>The workbook sheet entry with the name.</summary>
        private ExcelFormulaSheetEntry? FindFormulaSheetEntry(string sheetName) {
            foreach (ExcelFormulaSheetEntry entry in ReadFormulaSheetEntries()) {
                if (string.Equals(entry.Name, sheetName, StringComparison.OrdinalIgnoreCase)) return entry.IsWorksheet ? entry : null;
            }
            return null;
        }

        /// <summary>The tables of this sheet.</summary>
        private IEnumerable<ExcelFormulaTable> ReadFormulaTables() => FormulaModel.Tables();

        /// <summary>The tables of every worksheet in the workbook.</summary>
        private IEnumerable<ExcelFormulaTable> ReadWorkbookFormulaTables() {
            foreach ((_, ExcelFormulaTable table) in _excelDocument.FormulaWorkbookModel.Tables()) yield return table;
        }

        /// <summary>Finds the table with the name or display name, and the sheet entry that holds it.</summary>
        private bool TryFindFormulaTable(string tableName, out ExcelFormulaSheetEntry sheetEntry, out ExcelFormulaTable table) {
            foreach ((ExcelFormulaSheetEntry sheet, ExcelFormulaTable candidate) in _excelDocument.FormulaWorkbookModel.Tables()) {
                if (string.Equals(candidate.Name, tableName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(candidate.DisplayName, tableName, StringComparison.OrdinalIgnoreCase)) {
                    sheetEntry = sheet;
                    table = candidate;
                    return true;
                }
            }
            sheetEntry = null!;
            table = null!;
            return false;
        }

        private bool TryGetContainingFormulaTable(int row, int column, out ExcelFormulaTable table) {
            foreach (ExcelFormulaTable candidate in FormulaModel.Tables()) {
                if (candidate.Reference != null
                    && ExcelA1.TryParseRange(candidate.Reference.Replace("$", string.Empty),
                        out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
                    && row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn) {
                    table = candidate;
                    return true;
                }
            }
            table = null!;
            return false;
        }

        /// <summary>Marks every formula of the sheet to be calculated when the workbook opens.</summary>
        public void InvalidateFormulas() {
            WriteLock(MarkFormulaCellsForRecalculation);
        }

        internal void MarkFormulasForRecalculation() {
            MarkFormulaCellsForRecalculation();
            MarkRequiresSavePreparation();
        }

        private void MarkFormulaCellsForRecalculation() {
            foreach (ExcelFormulaCellRecord cell in FormulaModel.FormulaCells()) cell.SetCalculateCell(true);
        }

        /// <summary>
        /// Removes cached values from formula cells on this sheet.
        /// </summary>
        public void ClearCachedFormulaResults() {
            WriteLock(() => {
                ExcelFormulaSheetModel model = FormulaModel;
                List<ExcelFormulaCellRecord> formulas = new(model.FormulaCells());
                if (model.HasTableHeaders) {
                    foreach (ExcelFormulaCellRecord formula in formulas) {
                        if (FormulaTargetsTableHeader(new ExcelFormulaCell(formula)))
                            throw new InvalidOperationException("Clearing these formula caches would erase a table-column header.");
                    }
                }
                bool changed = false;
                var arrays = new List<(int Top, int Left, int Bottom, int Right, bool Value)>();
                foreach (ExcelFormulaCellRecord formula in formulas) {
                    if (formula.FormulaKind == ExcelFormulaKind.Array && formula.FormulaReference is string reference
                        && ExcelA1.TryParseRange(reference, out int top, out int left, out int bottom, out int right)) {
                        arrays.Add((top, left, bottom, right, true));
                    }
                }
                var arrayIndex = new ExcelCellRangeIndex<bool>(arrays);
                foreach (ExcelFormulaCellRecord cell in model.Cells().ToList()) {
                    bool arrayMember = !cell.HasFormula && arrays.Count != 0
                        && cell.Reference is string address
                        && TryParseCellReference(address, out int row, out int column)
                        && arrayIndex.At(row, column).Any();
                    if (!cell.HasFormula && !arrayMember) continue;
                    if (cell.ValueText != null || cell.ValueMetaIndex != null) {
                        cell.SetValue(null, cell.ValueType);
                        cell.ClearValueMetaIndex();
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
                List<ExcelFormulaCell> formulaCells = EnumerateFormulaCells().ToList();
                var dependencyInspectionContext = new FormulaDependencyInspectionContext(this, formulaCells, sharedFormulaDefinitions);
                ExcelFormulaWorkbookModel workbook = _excelDocument.FormulaWorkbookModel;
                ExcelCellMetadataIndex? metadata = workbook.CellMetadata;
                bool packageRequestsRecalculation = workbook.RequestsFullCalculation(this);
                Uri sheetKey = FormulaSheetKey;
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
                    if (workbook.HasCellMetadata) PlanDynamicArrayOwners();
                    foreach (ExcelFormulaCell cell in formulaCells) {
                        _formulaEvaluationGuardState.DependencyGuardBlocked = false;
                        string formula = ResolveCellFormulaText(cell, sharedFormulaDefinitions);
                        string cellReference = cell.Reference ?? string.Empty;
                        bool supported = TryEvaluateFormulaCellValue(cell, out _, sharedFormulaDefinitions);
                        IReadOnlyList<string> dependencies = GetFormulaDependencies(cell.Reference, formula, dependencyAliases, dependencyTables);
                        IReadOnlyList<string> dependencyIssues = GetFormulaDependencyIssues(cell.Reference, dependencies, dependencyInspectionContext);
                        formulas.Add(new ExcelFormulaCellInfo(
                            Name,
                            cellReference,
                            formula,
                            GetFormulaCachedText(cell),
                            packageRequestsRecalculation
                                || cell.CalculateCell == true
                                || _excelDocument.HasFormulaInputMutationsAfterFormulaBaseline(sheetKey, cellReference)
                                || HasFormulaDependencyMutationsAfterBaseline(cellReference, dependencies),
                            supported,
                            supported ? null : GetUnsupportedFormulaReason(formula),
                            dependencies,
                            dependencyIssues,
                            ExcelCellMetadataIndex.ArrayInfo(cell, metadata)));
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

        private bool HasFormulaDependencyMutationsAfterBaseline(string cellReference, IReadOnlyList<string> dependencies) {
            if (dependencies.Count == 0) return false;
            long baseline = _excelDocument.GetFormulaDependencyBaseline(FormulaSheetKey, cellReference);
            foreach (string dependency in dependencies) {
                if (!TryResolveFormulaDependencyReference(dependency, out ExcelSheet dependencySheet,
                    out int firstRow, out int firstColumn, out int lastRow, out int lastColumn, out _)) {
                    continue;
                }
                if (_excelDocument.HasFormulaDependencyMutationAfter(dependencySheet.FormulaSheetKey,
                    firstRow, firstColumn, lastRow, lastColumn, baseline)) {
                    return true;
                }
            }
            return false;
        }

        /// <summary>The formula texts of the sheet's shared formula members, resolved; throws when one cannot be.</summary>
        private IReadOnlyList<string> ResolveSharedFormulaTextsForStructuralValidation() {
            List<ExcelFormulaCell> sharedCells = FindSharedFormulaCells();
            if (sharedCells.Count == 0) return Array.Empty<string>();
            IReadOnlyDictionary<ExcelFormulaCell, (int Row, int Column)> effectiveCoordinates = BuildEffectiveCellCoordinates();
            IReadOnlyDictionary<uint, SharedFormulaDefinition> definitions = BuildSharedFormulaDefinitions(effectiveCoordinates);
            var resolved = new List<string>();
            foreach (ExcelFormulaCell cell in sharedCells) {
                string text = ResolveCellFormulaText(cell, definitions, effectiveCoordinates);
                if (string.IsNullOrWhiteSpace(text)) {
                    throw new InvalidOperationException(
                        $"Cannot edit rows because shared formula group {cell.SharedIndex!.Value} cannot be validated safely.");
                }
                resolved.Add(text);
            }
            return resolved;
        }

        private List<ExcelFormulaCell> FindSharedFormulaCells() {
            var cells = new List<ExcelFormulaCell>();
            foreach (ExcelFormulaCell cell in EnumerateFormulaCells()) {
                if (cell.FormulaKind == ExcelFormulaKind.Shared && cell.SharedIndex != null) cells.Add(cell);
            }
            return cells;
        }

        /// <summary>Gives every shared formula member its own formula text, which a structural edit can then rewrite.</summary>
        private void MaterializeSharedFormulasForStructuralEdit() {
            List<ExcelFormulaCell> sharedCells = FindSharedFormulaCells();
            if (sharedCells.Count == 0) return;
            IReadOnlyDictionary<ExcelFormulaCell, (int Row, int Column)> effectiveCoordinates = BuildEffectiveCellCoordinates();
            IReadOnlyDictionary<uint, SharedFormulaDefinition> definitions = BuildSharedFormulaDefinitions(effectiveCoordinates);
            var resolved = new List<(ExcelFormulaCellRecord Cell, string Text)>();
            foreach (ExcelFormulaCell cell in sharedCells) {
                if (cell.SharedIndex is not uint sharedIndex) continue;
                string text = ResolveCellFormulaText(cell, definitions, effectiveCoordinates);
                if (string.IsNullOrWhiteSpace(text)) {
                    throw new InvalidOperationException(
                        $"Cannot edit rows because shared formula group {sharedIndex} cannot be materialized safely.");
                }
                resolved.Add((cell.Record!, text));
            }
            foreach ((ExcelFormulaCellRecord cell, string text) in resolved) cell.SetOwnFormula(text);
        }
    }
}
