using System;
using System.Collections.Generic;
using System.Linq;

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

            foreach (ExcelSheet sourceSheet in sourceSheets) {
                string requestedName = (options.SheetNamePrefix ?? string.Empty) + sourceSheet.Name;
                ExcelSheet targetSheet = CopyWorkSheetFrom(sourceDocument, sourceSheet.Name, requestedName, options.SheetNameValidationMode, new ExcelWorksheetCopyOptions {
                    CopyMode = options.CopyMode
                });
                importedSourceNames.Add(sourceSheet.Name);
                createdTargetNames.Add(targetSheet.Name);
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
    }
}
