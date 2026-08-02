using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private long _formulaInputMutationVersion;
        private long _formulaAuthoredMutationVersion;
        private readonly ConcurrentDictionary<Uri, long> _formulaRecalculationVersions = new();
        private readonly ConcurrentDictionary<Uri, long> _formulaAuthoredRecalculationVersions = new();
        private readonly ConcurrentDictionary<(Uri WorksheetUri, string CellReference), long> _formulaAuthoredVersions = new();
        private readonly ConcurrentDictionary<(Uri WorksheetUri, string CellReference), long> _formulaCellRecalculationVersions = new();
        private readonly ConcurrentDictionary<(Uri WorksheetUri, string CellReference), long> _formulaDependencyBaselines = new();
        private readonly ConcurrentDictionary<(Uri WorksheetUri, string CellReference), long> _formulaCellDependencyRecalculationVersions = new();
        private readonly ConcurrentDictionary<(Uri WorksheetUri, string CellReference), long> _formulaDependencyMutationVersions = new();

        /// <summary>
        /// Formula calculation and cached-result policy used during save.
        /// </summary>
        public ExcelCalculationOptions Calculation { get; } = new ExcelCalculationOptions();

        internal void MarkFormulaInputMutation() {
            Interlocked.Increment(ref _formulaInputMutationVersion);
        }

        internal long CaptureFormulaInputMutationVersion() =>
            Interlocked.Read(ref _formulaInputMutationVersion);

        internal bool HasFormulaInputMutationsAfterLastRecalculation(WorksheetPart worksheetPart) {
            long mutationVersion = Interlocked.Read(ref _formulaInputMutationVersion);
            return mutationVersion > 0
                && (!_formulaRecalculationVersions.TryGetValue(worksheetPart.Uri, out long recalculationVersion)
                    || recalculationVersion < mutationVersion);
        }

        internal void MarkFormulaAuthored(
            WorksheetPart worksheetPart,
            string cellReference,
            bool retainedCachedValue = false) {
            var key = (worksheetPart.Uri, cellReference);
            _formulaAuthoredVersions[key] = Interlocked.Read(ref _formulaInputMutationVersion);
            long dependencyVersion = Interlocked.Increment(ref _formulaAuthoredMutationVersion);
            _formulaDependencyBaselines[key] = dependencyVersion;
            _formulaDependencyMutationVersions[key] = dependencyVersion;
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
            if (_formulaAuthoredVersions.TryGetValue((worksheetPart.Uri, cellReference), out long authoredVersion)
                && authoredVersion > baseline) {
                baseline = authoredVersion;
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
            if (_formulaDependencyBaselines.TryGetValue((worksheetPart.Uri, cellReference), out long authoredVersion)
                && authoredVersion > baseline) {
                baseline = authoredVersion;
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
                        if (_formulaDependencyMutationVersions.TryGetValue(
                                (worksheetPart.Uri, A1.CellReference(row, column)),
                                out long mutationVersion)
                            && mutationVersion > baseline) {
                            return true;
                        }
                    }
                }
                return false;
            }

            foreach (KeyValuePair<(Uri WorksheetUri, string CellReference), long> item in _formulaDependencyMutationVersions) {
                if (item.Value <= baseline
                    || item.Key.WorksheetUri != worksheetPart.Uri
                    || !A1.TryParseCellReferenceFast(item.Key.CellReference, out int row, out int column)) {
                    continue;
                }
                if (row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
            return false;
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
        /// Removes workbook-level structure/window protection metadata.
        /// </summary>
        public void UnprotectWorkbook() {
            var workbook = WorkbookRoot;
            var protection = workbook.GetFirstChild<WorkbookProtection>();
            if (protection != null) {
                workbook.RemoveChild(protection);
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
            ExcelSheet? pendingSheet = MaterializePendingDirectCellValueSheetIfNeeded();

            int count = 0;
            string? recalculatedPendingSheetName = null;
            if (pendingSheet != null) {
                count += pendingSheet.RecalculateSupportedFormulas();
                recalculatedPendingSheetName = pendingSheet.Name;
            }

            foreach (var sheet in Sheets) {
                if (recalculatedPendingSheetName != null
                    && string.Equals(sheet.Name, recalculatedPendingSheetName, System.StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                count += sheet.RecalculateSupportedFormulas();
            }

            return count;
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
