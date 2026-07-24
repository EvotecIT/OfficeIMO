using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents an Excel document and provides methods for creating,
    /// loading and saving spreadsheets.
    /// </summary>
    public partial class ExcelDocument {
        /// <summary>
        /// Imports selected or all worksheets from another workbook into this workbook.
        /// </summary>
        public ExcelWorkbookMergeResult MergeWorkbookFrom(ExcelDocument sourceDocument, ExcelWorkbookMergeOptions? options = null) {
            if (sourceDocument == null) {
                throw new ArgumentNullException(nameof(sourceDocument));
            }

            options ??= new ExcelWorkbookMergeOptions();
            if (options.MaxDefinedNames <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxDefinedNames));
            if (options.MaxDefinedNameCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxDefinedNameCharacters));
            var definedNameBudget = new DefinedNameCopyBudget(options.MaxDefinedNames, options.MaxDefinedNameCharacters);
            List<ExcelSheet> sourceSheets = ResolveWorkbookMergeSheets(sourceDocument, options).ToList();
            var importedSourceNames = new List<string>(sourceSheets.Count);
            var createdTargetNames = new List<string>(sourceSheets.Count);
            var sheetNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var tableNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var externalReferenceMaps = new Dictionary<string, IReadOnlyDictionary<int, int>>(StringComparer.OrdinalIgnoreCase);

            foreach (ExcelSheet sourceSheet in sourceSheets) {
                string requestedName = (options.SheetNamePrefix ?? string.Empty) + sourceSheet.Name;
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
                        definedNameBudget);
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
            for (int index = 0; index < importedSourceNames.Count; index++) {
                ExcelSheet targetSheet = GetSheet(createdTargetNames[index]);
                externalReferenceMaps.TryGetValue(targetSheet.Name, out IReadOnlyDictionary<int, int>? externalReferenceMap);
                CopyReferencedDefinedNamesFromSource(
                    sourceDocument,
                    targetSheet,
                    sheetNameMap,
                    tableNameMap,
                    externalReferenceMap,
                    definedNameBudget);
            }

            MarkPackageDirty();
            return new ExcelWorkbookMergeResult(importedSourceNames, createdTargetNames);
        }

        private static IEnumerable<ExcelSheet> ResolveWorkbookMergeSheets(ExcelDocument sourceDocument, ExcelWorkbookMergeOptions options) {
            if (options.SheetNames == null || options.SheetNames.Count == 0) {
                return sourceDocument.Sheets;
            }

            return options.SheetNames.Select(sourceDocument.GetSheet);
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
                RewriteCopiedWorksheetReferences(worksheetPart, sheetNameMap, tableNameMap);
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
