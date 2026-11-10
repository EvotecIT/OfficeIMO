using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private long _formulaInputMutationVersion;
        private long _formulaAuthoredMutationVersion;
        private readonly ConcurrentDictionary<Uri, long> _formulaRecalculationVersions = new();
        private readonly ConcurrentDictionary<Uri, long> _formulaAuthoredRecalculationVersions = new();
        // Authoring a formula records both versions at once, so one entry per cell holds them.
        private readonly FormulaCellVersionMap<FormulaAuthoredVersions> _formulaAuthoredVersions = new();
        private readonly FormulaCellVersionMap<long> _formulaCellRecalculationVersions = new();
        private readonly FormulaCellVersionMap<long> _formulaCellDependencyRecalculationVersions = new();

        /// <summary>The input-mutation version a formula was authored at, and its dependency version.</summary>
        internal readonly struct FormulaAuthoredVersions {
            internal FormulaAuthoredVersions(long input, long dependency) {
                Input = input;
                Dependency = dependency;
            }

            internal long Input { get; }
            internal long Dependency { get; }
        }

        /// <summary>
        /// Formula calculation and cached-result policy used during save.
        /// </summary>
        public ExcelCalculationOptions Calculation { get; } = new ExcelCalculationOptions();

        internal void MarkFormulaInputMutation() {
            Interlocked.Increment(ref _formulaInputMutationVersion);
        }

        internal long CaptureFormulaInputMutationVersion() =>
            Interlocked.Read(ref _formulaInputMutationVersion);

        internal FormulaMutationState CaptureFormulaMutationState() => new FormulaMutationState(
            Interlocked.Read(ref _formulaInputMutationVersion),
            Interlocked.Read(ref _formulaAuthoredMutationVersion),
            new Dictionary<Uri, long>(_formulaRecalculationVersions),
            new Dictionary<Uri, long>(_formulaAuthoredRecalculationVersions),
            _formulaAuthoredVersions.Snapshot(),
            _formulaCellRecalculationVersions.Snapshot(),
            _formulaCellDependencyRecalculationVersions.Snapshot());

        internal void RestoreFormulaMutationState(FormulaMutationState state) {
            ArgumentNullException.ThrowIfNull(state);
            Interlocked.Exchange(ref _formulaInputMutationVersion, state.InputMutationVersion);
            Interlocked.Exchange(ref _formulaAuthoredMutationVersion, state.AuthoredMutationVersion);
            RestoreFormulaMutationDictionary(_formulaRecalculationVersions, state.RecalculationVersions);
            RestoreFormulaMutationDictionary(_formulaAuthoredRecalculationVersions, state.AuthoredRecalculationVersions);
            _formulaAuthoredVersions.Restore(state.AuthoredVersions);
            _formulaCellRecalculationVersions.Restore(state.CellRecalculationVersions);
            _formulaCellDependencyRecalculationVersions.Restore(state.CellDependencyRecalculationVersions);
        }

        private static void RestoreFormulaMutationDictionary<TKey>(
            ConcurrentDictionary<TKey, long> target,
            IReadOnlyDictionary<TKey, long> snapshot) where TKey : notnull {
            target.Clear();
            foreach (KeyValuePair<TKey, long> item in snapshot) target[item.Key] = item.Value;
        }

        internal sealed class FormulaMutationState {
            internal FormulaMutationState(
                long inputMutationVersion,
                long authoredMutationVersion,
                IReadOnlyDictionary<Uri, long> recalculationVersions,
                IReadOnlyDictionary<Uri, long> authoredRecalculationVersions,
                IReadOnlyDictionary<(Uri WorksheetUri, string CellReference), FormulaAuthoredVersions> authoredVersions,
                IReadOnlyDictionary<(Uri WorksheetUri, string CellReference), long> cellRecalculationVersions,
                IReadOnlyDictionary<(Uri WorksheetUri, string CellReference), long> cellDependencyRecalculationVersions) {
                InputMutationVersion = inputMutationVersion;
                AuthoredMutationVersion = authoredMutationVersion;
                RecalculationVersions = recalculationVersions;
                AuthoredRecalculationVersions = authoredRecalculationVersions;
                AuthoredVersions = authoredVersions;
                CellRecalculationVersions = cellRecalculationVersions;
                CellDependencyRecalculationVersions = cellDependencyRecalculationVersions;
            }

            internal long InputMutationVersion { get; }
            internal long AuthoredMutationVersion { get; }
            internal IReadOnlyDictionary<Uri, long> RecalculationVersions { get; }
            internal IReadOnlyDictionary<Uri, long> AuthoredRecalculationVersions { get; }
            internal IReadOnlyDictionary<(Uri WorksheetUri, string CellReference), FormulaAuthoredVersions> AuthoredVersions { get; }
            internal IReadOnlyDictionary<(Uri WorksheetUri, string CellReference), long> CellRecalculationVersions { get; }
            internal IReadOnlyDictionary<(Uri WorksheetUri, string CellReference), long> CellDependencyRecalculationVersions { get; }
        }

        internal bool HasFormulaInputMutationsAfterLastRecalculation(WorksheetPart worksheetPart) {
            long mutationVersion = Interlocked.Read(ref _formulaInputMutationVersion);
            return mutationVersion > 0
                && (!_formulaRecalculationVersions.TryGetValue(worksheetPart.Uri, out long recalculationVersion)
                    || recalculationVersion < mutationVersion);
        }

        /// <summary>
        /// Per-cell formula versions. Each authored or recalculated formula adds entries, and a lock over one
        /// dictionary stores them in its entry array instead of allocating a node per entry as a concurrent map does.
        /// </summary>
        private sealed class FormulaCellVersionMap<TValue> where TValue : struct {
            private readonly Dictionary<(Uri WorksheetUri, string CellReference), TValue> _versions = new(FormulaCellKeyComparer.Instance);

            internal TValue this[(Uri WorksheetUri, string CellReference) key] {
                set {
                    lock (_versions) _versions[key] = value;
                }
            }

            internal bool TryGetValue((Uri WorksheetUri, string CellReference) key, out TValue value) {
                lock (_versions) return _versions.TryGetValue(key, out value);
            }

            internal Dictionary<(Uri WorksheetUri, string CellReference), TValue> Snapshot() {
                lock (_versions) return new Dictionary<(Uri WorksheetUri, string CellReference), TValue>(_versions, FormulaCellKeyComparer.Instance);
            }

            internal bool Any(Func<(Uri WorksheetUri, string CellReference), TValue, bool> predicate) {
                lock (_versions) {
                    foreach (KeyValuePair<(Uri WorksheetUri, string CellReference), TValue> item in _versions) {
                        if (predicate(item.Key, item.Value)) return true;
                    }
                }
                return false;
            }

            internal void Restore(IReadOnlyDictionary<(Uri WorksheetUri, string CellReference), TValue> snapshot) {
                lock (_versions) {
                    _versions.Clear();
                    foreach (KeyValuePair<(Uri WorksheetUri, string CellReference), TValue> item in snapshot) _versions[item.Key] = item.Value;
                }
            }
        }

        /// <summary>Hashes the cell text only, because hashing a part URI per formula dominated authoring.</summary>
        private sealed class FormulaCellKeyComparer : IEqualityComparer<(Uri WorksheetUri, string CellReference)> {
            internal static readonly FormulaCellKeyComparer Instance = new();

            public bool Equals((Uri WorksheetUri, string CellReference) left, (Uri WorksheetUri, string CellReference) right) =>
                string.Equals(left.CellReference, right.CellReference, StringComparison.Ordinal)
                && (ReferenceEquals(left.WorksheetUri, right.WorksheetUri) || left.WorksheetUri == right.WorksheetUri);

            public int GetHashCode((Uri WorksheetUri, string CellReference) key) => StringComparer.Ordinal.GetHashCode(key.CellReference);
        }

        internal void MarkFormulaAuthored(
            WorksheetPart worksheetPart,
            string cellReference,
            bool retainedCachedValue = false) {
            long inputVersion = Interlocked.Read(ref _formulaInputMutationVersion);
            long dependencyVersion = Interlocked.Increment(ref _formulaAuthoredMutationVersion);
            _formulaAuthoredVersions[(worksheetPart.Uri, cellReference)] = new FormulaAuthoredVersions(inputVersion, dependencyVersion);
            if (retainedCachedValue) MarkFormulaInputMutation();
        }

        internal bool HasFormulaInputMutationsAfterFormulaBaseline(WorksheetPart worksheetPart, string cellReference) {
            long mutationVersion = Interlocked.Read(ref _formulaInputMutationVersion);
            if (mutationVersion == 0) {
                return false;
            }

            long baseline = 0;
            if (_formulaRecalculationVersions.TryGetValue(worksheetPart.Uri, out long recalculationVersion)) {
                baseline = recalculationVersion;
            }
            if (_formulaAuthoredVersions.TryGetValue((worksheetPart.Uri, cellReference), out FormulaAuthoredVersions authored)
                && authored.Input > baseline) {
                baseline = authored.Input;
            }
            if (_formulaCellRecalculationVersions.TryGetValue((worksheetPart.Uri, cellReference), out long cellRecalculationVersion)
                && cellRecalculationVersion > baseline) {
                baseline = cellRecalculationVersion;
            }

            return mutationVersion > baseline;
        }

        internal void MarkFormulaSheetRecalculated(WorksheetPart worksheetPart, long mutationVersion) {
            _formulaRecalculationVersions[worksheetPart.Uri] = mutationVersion;
            _formulaAuthoredRecalculationVersions[worksheetPart.Uri] =
                Interlocked.Read(ref _formulaAuthoredMutationVersion);
        }

        internal void MarkFormulaCellRecalculated(
            WorksheetPart worksheetPart,
            string cellReference,
            long mutationVersion) {
            var key = (worksheetPart.Uri, cellReference);
            _formulaCellRecalculationVersions[key] = mutationVersion;
            _formulaCellDependencyRecalculationVersions[key] =
                Interlocked.Read(ref _formulaAuthoredMutationVersion);
        }

        internal long GetFormulaDependencyBaseline(WorksheetPart worksheetPart, string cellReference) {
            long baseline = 0;
            if (_formulaAuthoredRecalculationVersions.TryGetValue(worksheetPart.Uri, out long recalculationVersion)) {
                baseline = recalculationVersion;
            }
            if (_formulaAuthoredVersions.TryGetValue((worksheetPart.Uri, cellReference), out FormulaAuthoredVersions authored)
                && authored.Dependency > baseline) {
                baseline = authored.Dependency;
            }
            if (_formulaCellDependencyRecalculationVersions.TryGetValue(
                    (worksheetPart.Uri, cellReference),
                    out long cellRecalculationVersion)
                && cellRecalculationVersion > baseline) {
                baseline = cellRecalculationVersion;
            }
            return baseline;
        }

        internal bool HasFormulaDependencyMutationAfter(
            WorksheetPart worksheetPart,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            long baseline) {
            long cellCount = (long)(lastRow - firstRow + 1) * (lastColumn - firstColumn + 1);
            if (cellCount <= 4096L) {
                for (int row = firstRow; row <= lastRow; row++) {
                    for (int column = firstColumn; column <= lastColumn; column++) {
                        if (_formulaAuthoredVersions.TryGetValue(
                                (worksheetPart.Uri, ExcelA1.CellReference(row, column)),
                                out FormulaAuthoredVersions authored)
                            && authored.Dependency > baseline) {
                            return true;
                        }
                    }
                }
                return false;
            }

            return _formulaAuthoredVersions.Any((key, version) =>
                version.Dependency > baseline
                && key.WorksheetUri == worksheetPart.Uri
                && ExcelA1.TryParseCellReferenceFast(key.CellReference, out int row, out int column)
                && row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn);
        }

        /// <summary>
        /// Returns true when workbook-level structure or window protection is present.
        /// </summary>
        public bool IsWorkbookProtected {
            get {
                var protection = WorkbookRoot.GetFirstChild<WorkbookProtection>();
                return protection != null && ((protection.LockStructure?.Value ?? false) || (protection.LockWindows?.Value ?? false));
            }
        }

        /// <summary>
        /// Protects workbook structure/window metadata. This is not file encryption.
        /// </summary>
        public void ProtectWorkbook(ExcelWorkbookProtectionOptions? options = null) {
            var opts = options ?? new ExcelWorkbookProtectionOptions();
            var workbook = WorkbookRoot;
            var protection = workbook.GetFirstChild<WorkbookProtection>();
            if (protection == null) {
                protection = new WorkbookProtection();
                var workbookViews = workbook.GetFirstChild<BookViews>();
                if (workbookViews != null) {
                    workbook.InsertBefore(protection, workbookViews);
                } else if (workbook.GetFirstChild<Sheets>() is Sheets sheets) {
                    workbook.InsertBefore(protection, sheets);
                } else if (workbook.GetFirstChild<WorkbookProperties>() is WorkbookProperties workbookProperties) {
                    workbook.InsertAfter(protection, workbookProperties);
                } else if (workbook.GetFirstChild<FileSharing>() is FileSharing fileSharing) {
                    workbook.InsertAfter(protection, fileSharing);
                } else if (workbook.GetFirstChild<FileVersion>() is FileVersion fileVersion) {
                    workbook.InsertAfter(protection, fileVersion);
                } else {
                    workbook.InsertAt(protection, 0);
                }
            }

            protection.LockStructure = opts.ProtectStructure;
            protection.LockWindows = opts.ProtectWindows;
            string? hash = ExcelProtectionHash.ResolveLegacyHash(opts.Password, opts.LegacyPasswordHash);
            if (hash != null) {
                protection.WorkbookPassword = hash;
            } else {
                protection.WorkbookPassword = null;
                protection.RemoveAttribute("workbookPassword", string.Empty);
            }
            workbook.Save();
            MarkPackageDirty();
        }

        /// <summary>
        /// Removes the selected workbook structure/window restrictions. With both selected (the default),
        /// removes all workbook protection metadata. A subset preserves other restrictions and password hashes.
        /// </summary>
        /// <param name="structure">Whether to remove the workbook structure restriction.</param>
        /// <param name="windows">Whether to remove the workbook window restriction.</param>
        public void UnprotectWorkbook(bool structure = true, bool windows = true) {
            if (!structure && !windows) {
                throw new ArgumentException("Select at least one workbook restriction to remove.");
            }
            var workbook = WorkbookRoot;
            var protection = workbook.GetFirstChild<WorkbookProtection>();
            if (protection != null) {
                if (structure && windows) {
                    workbook.RemoveChild(protection);
                } else {
                    if (structure) protection.LockStructure = false;
                    if (windows) protection.LockWindows = false;
                    if (protection.LockStructure?.Value != true && protection.LockWindows?.Value != true && protection.LockRevision?.Value != true) {
                        workbook.RemoveChild(protection);
                    }
                }
                workbook.Save();
                MarkPackageDirty();
            }
        }

        /// <summary>
        /// Marks all formulas dirty so Excel-compatible applications recalculate them on open.
        /// </summary>
        public void InvalidateFormulas() {
            foreach (var sheet in Sheets) {
                sheet.InvalidateFormulas();
            }

            ConfigureFullCalculationOnOpen();
        }

        /// <summary>
        /// Removes cached values from all formula cells.
        /// </summary>
        public void ClearCachedFormulaResults() {
            foreach (var sheet in Sheets) {
                sheet.ClearCachedFormulaResults();
            }
        }

        /// <summary>
        /// Evaluates supported formulas and writes cached values.
        /// </summary>
        public int RecalculateSupportedFormulas() {
            return Locking.ExecuteWrite(EnsureLock(), () => {
                ExcelSheet? pendingSheet = MaterializePendingDirectCellValueSheetIfNeeded();
                using var context = new ExcelSheet.FormulaCalculationContext();
                int count = 0;
                if (pendingSheet != null) count += pendingSheet.RecalculateSupportedFormulas(context);
                foreach (var sheet in GetSheetsForLockedOperation()) {
                    if (pendingSheet != null && string.Equals(sheet.Name, pendingSheet.Name, StringComparison.OrdinalIgnoreCase)) continue;
                    count += sheet.RecalculateSupportedFormulas(context);
                }
                return count;
            });
        }

        /// <summary>
        /// Calculates formulas that are supported by OfficeIMO's lightweight formula engine and writes cached values.
        /// Unsupported formulas are preserved unchanged for Excel-compatible applications to calculate.
        /// </summary>
        /// <returns>The number of formula cells with updated cached values.</returns>
        public int Calculate() {
            return RecalculateSupportedFormulas();
        }

        /// <summary>
        /// Inspects formula cells across all worksheets without changing workbook contents.
        /// </summary>
        public ExcelFormulaInspection InspectFormulas() {
            var formulas = new List<ExcelFormulaCellInfo>();
            var workbookPart = WorkbookPartRoot;
            foreach (Sheet sheetElement in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (string.IsNullOrWhiteSpace(sheetElement.Id?.Value)) {
                    continue;
                }

                if (workbookPart.GetPartById(sheetElement.Id!.Value!) is not WorksheetPart) {
                    continue;
                }

                var sheet = new ExcelSheet(this, _spreadSheetDocument!, sheetElement);
                formulas.AddRange(sheet.GetFormulaCells());
            }

            return new ExcelFormulaInspection(formulas);
        }

        /// <summary>
        /// Requests a full workbook recalculation when the file is opened.
        /// </summary>
        public void ConfigureFullCalculationOnOpen() {
            var workbook = WorkbookRoot;
            var properties = workbook.GetFirstChild<CalculationProperties>();
            if (properties == null) {
                properties = new CalculationProperties();
            } else {
                properties.Remove();
            }

            InsertCalculationPropertiesInSchemaOrder(workbook, properties);
            properties.ForceFullCalculation = true;
            properties.FullCalculationOnLoad = true;
            workbook.Save();
            MarkPackageDirty();
        }

        private static void InsertCalculationPropertiesInSchemaOrder(Workbook workbook, CalculationProperties properties) {
            var laterChild = workbook.ChildElements.FirstOrDefault(child =>
                string.Equals(child.LocalName, "oleSize", StringComparison.Ordinal)
                || string.Equals(child.LocalName, "customWorkbookViews", StringComparison.Ordinal)
                || string.Equals(child.LocalName, "pivotCaches", StringComparison.Ordinal)
                || string.Equals(child.LocalName, "smartTagPr", StringComparison.Ordinal)
                || string.Equals(child.LocalName, "smartTagTypes", StringComparison.Ordinal)
                || string.Equals(child.LocalName, "webPublishing", StringComparison.Ordinal)
                || string.Equals(child.LocalName, "fileRecoveryPr", StringComparison.Ordinal)
                || string.Equals(child.LocalName, "webPublishObjects", StringComparison.Ordinal)
                || string.Equals(child.LocalName, "extLst", StringComparison.Ordinal));

            if (laterChild != null) {
                workbook.InsertBefore(properties, laterChild);
            } else {
                workbook.Append(properties);
            }
        }

        internal void ApplyCalculationPolicyBeforeSave(ExcelSaveOptions? options) {
            if (ShouldEvaluateFormulasBeforeSave(options)) {
                RecalculateSupportedFormulas();
            } else if (ShouldClearCachedFormulaResultsBeforeSave(options)) {
                ClearCachedFormulaResults();
            }

            if (ShouldMarkFormulasDirtyBeforeSave(options)) {
                InvalidateFormulas();
            }

            if (ShouldForceFullCalculationOnOpen(options)) {
                ConfigureFullCalculationOnOpen();
            }
        }

        private bool ShouldEvaluateFormulasBeforeSave(ExcelSaveOptions? options) {
            return Calculation.EvaluateFormulasBeforeSave || options?.EvaluateFormulasBeforeSave == true;
        }

        private bool ShouldClearCachedFormulaResultsBeforeSave(ExcelSaveOptions? options) {
            return Calculation.ClearCachedFormulaResultsBeforeSave || options?.ClearCachedFormulaResultsBeforeSave == true;
        }

        private bool ShouldMarkFormulasDirtyBeforeSave(ExcelSaveOptions? options) {
            return Calculation.MarkFormulasDirtyBeforeSave || options?.MarkFormulasDirtyBeforeSave == true;
        }

        private bool ShouldForceFullCalculationOnOpen(ExcelSaveOptions? options) {
            return Calculation.ForceFullCalculationOnOpen || options?.ForceFullCalculationOnOpen == true;
        }

        private bool HasCalculationSaveWork(ExcelSaveOptions? options) {
            return ShouldEvaluateFormulasBeforeSave(options)
                || ShouldClearCachedFormulaResultsBeforeSave(options)
                || ShouldMarkFormulasDirtyBeforeSave(options)
                || ShouldForceFullCalculationOnOpen(options);
        }
    }
}
