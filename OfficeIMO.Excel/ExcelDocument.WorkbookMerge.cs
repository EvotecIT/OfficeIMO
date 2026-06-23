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
            List<ExcelSheet> sourceSheets = ResolveWorkbookMergeSheets(sourceDocument, options).ToList();
            var importedSourceNames = new List<string>(sourceSheets.Count);
            var createdTargetNames = new List<string>(sourceSheets.Count);
            var sheetNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (ExcelSheet sourceSheet in sourceSheets) {
                string requestedName = (options.SheetNamePrefix ?? string.Empty) + sourceSheet.Name;
                ExcelSheet targetSheet = CopyWorkSheetFrom(sourceDocument, sourceSheet.Name, requestedName, options.SheetNameValidationMode, new ExcelWorksheetCopyOptions {
                    CopyMode = options.CopyMode
                });
                importedSourceNames.Add(sourceSheet.Name);
                createdTargetNames.Add(targetSheet.Name);
                sheetNameMap[sourceSheet.Name] = targetSheet.Name;
            }

            RewriteMergedWorksheetReferences(createdTargetNames, sheetNameMap);
            for (int index = 0; index < importedSourceNames.Count; index++) {
                CopyReferencedDefinedNamesFromSource(sourceDocument, importedSourceNames[index], createdTargetNames[index], sheetNameMap);
            }

            MarkPackageDirty();
            return new ExcelWorkbookMergeResult(importedSourceNames, createdTargetNames);
        }

        /// <summary>
        /// Alias for <see cref="MergeWorkbookFrom(ExcelDocument, ExcelWorkbookMergeOptions?)"/>.
        /// </summary>
        public ExcelWorkbookMergeResult JoinWorkbookFrom(ExcelDocument sourceDocument, ExcelWorkbookMergeOptions? options = null)
            => MergeWorkbookFrom(sourceDocument, options);

        private static IEnumerable<ExcelSheet> ResolveWorkbookMergeSheets(ExcelDocument sourceDocument, ExcelWorkbookMergeOptions options) {
            if (options.SheetNames == null || options.SheetNames.Count == 0) {
                return sourceDocument.Sheets;
            }

            return options.SheetNames.Select(sourceDocument.GetSheet);
        }

        private void RewriteMergedWorksheetReferences(IEnumerable<string> copiedSheetNames, IReadOnlyDictionary<string, string> sheetNameMap) {
            if (sheetNameMap.Count == 0) {
                return;
            }

            foreach (string copiedSheetName in copiedSheetNames) {
                ExcelSheet copiedSheet = GetSheet(copiedSheetName);
                WorksheetPart worksheetPart = copiedSheet.WorksheetPart;
                RewriteCopiedWorksheetReferences(worksheetPart, sheetNameMap);
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
