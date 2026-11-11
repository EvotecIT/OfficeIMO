using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;

namespace OfficeIMO.Excel {
    // Represents an Excel document and provides methods for creating,
    // loading and saving spreadsheets.
    public partial class ExcelDocument {
        /// <summary>
        /// Imports selected or all worksheets from another workbook into this workbook.
        /// </summary>
        /// <remarks>
        /// Imported worksheets are named by <see cref="ExcelWorkbookMergeOptions.TargetSheetNames"/> or
        /// <see cref="ExcelWorkbookMergeOptions.SheetNamePrefix"/>. References between imported worksheets are
        /// rewritten from each source name to its target name in one step. Imported worksheets keep no tab
        /// selection from their source workbook, so only the active worksheet stays selected.
        /// Imported worksheets keep their source's hidden or very hidden state. When the merge leaves the active
        /// worksheet hidden, as when an empty workbook imports a hidden sheet first, the first visible worksheet becomes
        /// active. When the merge leaves no sheet visible, the first imported worksheet is shown.
        /// </remarks>
        public ExcelWorkbookMergeResult MergeWorkbookFrom(ExcelDocument sourceDocument, ExcelWorkbookMergeOptions? options = null) {
            ArgumentNullException.ThrowIfNull(sourceDocument);

            options ??= new ExcelWorkbookMergeOptions();
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxDefinedNames, nameof(options.MaxDefinedNames));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxDefinedNameCharacters, nameof(options.MaxDefinedNameCharacters));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxInCellImages, nameof(options.MaxInCellImages));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxInCellImageBytes, nameof(options.MaxInCellImageBytes));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxTotalInCellImageBytes, nameof(options.MaxTotalInCellImageBytes));
            var definedNameBudget = new DefinedNameCopyBudget(options.MaxDefinedNames, options.MaxDefinedNameCharacters);
            var inCellImageContext = new InCellImageCopyContext(
                options.MaxInCellImages,
                options.MaxInCellImageBytes,
                options.MaxTotalInCellImageBytes);
            List<ExcelSheet> sourceSheets = ResolveWorkbookMergeSheets(sourceDocument, options).ToList();
            Dictionary<string, string>? targetSheetNames = ResolveWorkbookMergeTargetSheetNames(sourceSheets, options);
            if (options.CopyMode == ExcelWorksheetCopyMode.Package) {
                foreach (ExcelSheet sourceSheet in sourceSheets) {
                    sourceSheet.PreflightInCellImages(inCellImageContext);
                }
                if (!ReferenceEquals(sourceDocument, this)) {
                    PreflightReferencedDefinedNamesFromSource(
                        sourceDocument,
                        sourceSheets,
                        definedNameBudget);
                }
            }
            var importedSourceNames = new List<string>(sourceSheets.Count);
            var createdTargetNames = new List<string>(sourceSheets.Count);
            var sheetNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var tableNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var externalReferenceMaps = new Dictionary<string, IReadOnlyDictionary<int, int>>(StringComparer.OrdinalIgnoreCase);

            foreach (ExcelSheet sourceSheet in sourceSheets) {
                string requestedName = ResolveWorkbookMergeSheetName(sourceSheet, options, targetSheetNames);
                ExcelSheet targetSheet;
                if (options.CopyMode == ExcelWorksheetCopyMode.Values) {
                    targetSheet = CopyWorksheetFromValues(sourceDocument, sourceSheet.Name, requestedName, options.SheetNameValidationMode);
                } else if (ReferenceEquals(sourceDocument, this)) {
                    WorksheetPackageCopyResult copyResult = CopyWorksheetWithinWorkbook(sourceSheet, requestedName, options.SheetNameValidationMode);
                    targetSheet = copyResult.Sheet;
                    foreach (var tableName in copyResult.TableNameMap) {
                        tableNameMap[tableName.Key] = tableName.Value;
                    }
                } else {
                    WorksheetPackageCopyResult copyResult = CopyWorksheetFromPackage(
                        sourceDocument,
                        sourceSheet.Name,
                        requestedName,
                        options.SheetNameValidationMode,
                        rewriteCopiedReferences: false,
                        copyReferencedDefinedNames: false,
                        options.CopyExternalWorkbookReferences,
                        inCellImageContext);
                    targetSheet = copyResult.Sheet;
                    foreach (var tableName in copyResult.TableNameMap) {
                        tableNameMap[tableName.Key] = tableName.Value;
                    }

                    if (copyResult.ExternalReferenceMap.Count > 0) {
                        externalReferenceMaps[targetSheet.Name] = copyResult.ExternalReferenceMap;
                    }
                }

                importedSourceNames.Add(sourceSheet.Name);
                createdTargetNames.Add(targetSheet.Name);
                sheetNameMap[sourceSheet.Name] = targetSheet.Name;
            }

            RewriteMergedWorksheetReferences(createdTargetNames, sheetNameMap, tableNameMap);
            if (options.CopyMode == ExcelWorksheetCopyMode.Package) {
                for (int index = 0; index < importedSourceNames.Count; index++) {
                    ExcelSheet targetSheet = GetSheet(createdTargetNames[index]);
                    externalReferenceMaps.TryGetValue(targetSheet.Name, out IReadOnlyDictionary<int, int>? externalReferenceMap);
                    CopyReferencedDefinedNamesFromSource(
                        sourceDocument,
                        targetSheet,
                        sheetNameMap,
                        tableNameMap,
                        externalReferenceMap);
                }
            }

            if (createdTargetNames.Count > 0) {
                EnsureVisibleActiveWorksheetAfterCopy(GetSheet(createdTargetNames[0]));
            }

            ClearMergedWorksheetTabSelection(createdTargetNames);
            MarkPackageDirty();
            return new ExcelWorkbookMergeResult(importedSourceNames, createdTargetNames);
        }

        private static IEnumerable<ExcelSheet> ResolveWorkbookMergeSheets(ExcelDocument sourceDocument, ExcelWorkbookMergeOptions options) {
            if (options.SheetNames == null || options.SheetNames.Count == 0) {
                return sourceDocument.Sheets;
            }

            return options.SheetNames.Select(sourceDocument.GetSheet);
        }

        private static string ResolveWorkbookMergeSheetName(
            ExcelSheet sourceSheet,
            ExcelWorkbookMergeOptions options,
            IReadOnlyDictionary<string, string>? targetSheetNames) {
            if (targetSheetNames != null && targetSheetNames.TryGetValue(sourceSheet.Name, out string? targetSheetName)) {
                return targetSheetName!;
            }

            return (options.SheetNamePrefix ?? string.Empty) + sourceSheet.Name;
        }

        private Dictionary<string, string>? ResolveWorkbookMergeTargetSheetNames(
            IReadOnlyList<ExcelSheet> sourceSheets,
            ExcelWorkbookMergeOptions options) {
            if (options.TargetSheetNames == null || options.TargetSheetNames.Count == 0) {
                return null;
            }

            var importedSourceNames = new HashSet<string>(sourceSheets.Select(_ => _.Name), StringComparer.OrdinalIgnoreCase);
            var targetSheetNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, string> entry in options.TargetSheetNames) {
                if (entry.Key == null || !importedSourceNames.Contains(entry.Key)) {
                    throw new ArgumentException(
                        $"TargetSheetNames names source worksheet '{entry.Key}', which this merge does not import.",
                        nameof(options));
                }

                if (entry.Value == null) {
                    throw new ArgumentException(
                        $"TargetSheetNames has no target name for source worksheet '{entry.Key}'.",
                        nameof(options));
                }

                if (targetSheetNames.ContainsKey(entry.Key)) {
                    throw new ArgumentException(
                        $"TargetSheetNames names source worksheet '{entry.Key}' more than once.",
                        nameof(options));
                }

                targetSheetNames.Add(entry.Key, entry.Value);
            }

            if (options.SheetNameValidationMode == ExcelSheetNameValidationMode.Strict) {
                PreflightStrictWorkbookMergeSheetNames(sourceSheets, options, targetSheetNames);
            }

            return targetSheetNames;
        }

        /// <summary>
        /// Applies the strict per-sheet validation to every requested name, and checks the names against each other,
        /// so an invalid or colliding name fails the merge before any worksheet is imported.
        /// </summary>
        private void PreflightStrictWorkbookMergeSheetNames(
            IReadOnlyList<ExcelSheet> sourceSheets,
            ExcelWorkbookMergeOptions options,
            IReadOnlyDictionary<string, string> targetSheetNames) {
            Locking.ExecuteRead(EnsureLock(), () => {
                var plannedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (ExcelSheet sourceSheet in sourceSheets) {
                    string requestedName = ResolveWorkbookMergeSheetName(sourceSheet, options, targetSheetNames);
                    string validatedName = ValidateOrSanitizeSheetName(requestedName, ExcelSheetNameValidationMode.Strict, currentSheetName: null);
                    if (!plannedNames.Add(validatedName)) {
                        throw new ArgumentException($"Worksheet name '{validatedName}' is requested for more than one imported worksheet.", nameof(options));
                    }
                }
            });
        }

        /// <summary>
        /// Copied worksheet XML keeps the source workbook's tab selection, which makes spreadsheet applications open
        /// the merged workbook with tabs grouped. Only the active worksheet keeps its selection.
        /// </summary>
        private void ClearMergedWorksheetTabSelection(IReadOnlyList<string> createdTargetNames) {
            string? activeSheetName = ActiveWorksheet?.Name;
            foreach (string createdTargetName in createdTargetNames) {
                if (string.Equals(createdTargetName, activeSheetName, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                Worksheet worksheet = GetSheet(createdTargetName).WorksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
                bool changed = false;
                foreach (SheetView sheetView in worksheet.Elements<SheetViews>().SelectMany(_ => _.Elements<SheetView>())) {
                    if (sheetView.TabSelected?.Value == true) {
                        sheetView.TabSelected = null;
                        changed = true;
                    }
                }

                if (changed) {
                    worksheet.Save();
                }
            }
        }

        private void RewriteMergedWorksheetReferences(
            IEnumerable<string> copiedSheetNames,
            IReadOnlyDictionary<string, string> sheetNameMap,
            IReadOnlyDictionary<string, string> tableNameMap) {
            if (sheetNameMap.Count == 0 && tableNameMap.Count == 0) {
                return;
            }

            foreach (string copiedSheetName in copiedSheetNames) {
                ExcelSheet copiedSheet = GetSheet(copiedSheetName);
                WorksheetPart worksheetPart = copiedSheet.WorksheetPart;
                RewriteCopiedWorksheetReferences(worksheetPart, sheetNameMap, tableNameMap, copiedSheet.Name);
            }
        }

        private static bool RewriteWorksheetSheetReferences(Worksheet worksheet, IReadOnlyDictionary<string, string> sheetNameMap) {
            bool changed = false;
            foreach (CellFormula formula in worksheet.Descendants<CellFormula>()) {
                changed |= RewriteFormulaSheetReference(formula, sheetNameMap);
            }

            foreach (Formula formula in worksheet.Descendants<Formula>()) {
                changed |= RewriteFormulaSheetReference(formula, sheetNameMap);
            }

            foreach (Formula1 formula in worksheet.Descendants<Formula1>()) {
                changed |= RewriteFormulaSheetReference(formula, sheetNameMap);
            }

            foreach (Formula2 formula in worksheet.Descendants<Formula2>()) {
                changed |= RewriteFormulaSheetReference(formula, sheetNameMap);
            }

            foreach (OfficeFormula formula in worksheet.Descendants<OfficeFormula>()) {
                changed |= RewriteFormulaSheetReference(formula, sheetNameMap);
            }

            foreach (Hyperlink hyperlink in worksheet.Descendants<Hyperlink>()) {
                // A relationship-backed bookmark belongs to its target document.
                if (!string.IsNullOrEmpty(hyperlink.Id?.Value)) continue;
                string? location = hyperlink.Location?.Value;
                if (string.IsNullOrEmpty(location)) {
                    continue;
                }

                string updated = ReplaceSheetNameReferences(location!, sheetNameMap);
                if (!string.Equals(updated, location, StringComparison.Ordinal)) {
                    hyperlink.Location = updated;
                    changed = true;
                }
            }

            return changed;
        }

        private static bool RewriteFormulaSheetReference(OpenXmlLeafTextElement formula, IReadOnlyDictionary<string, string> sheetNameMap) {
            string? text = formula.Text;
            if (string.IsNullOrEmpty(text)) {
                return false;
            }

            string updated = ReplaceSheetNameReferences(text!, sheetNameMap);

            if (string.Equals(updated, text, StringComparison.Ordinal)) {
                return false;
            }

            formula.Text = updated;
            return true;
        }
    }
}
