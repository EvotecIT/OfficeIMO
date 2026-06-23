using System.Collections.Generic;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Result of importing worksheets from one workbook into another.
    /// </summary>
    public sealed class ExcelWorkbookMergeResult {
        /// <summary>
        /// Creates a workbook merge result.
        /// </summary>
        public ExcelWorkbookMergeResult(IReadOnlyList<string> sourceSheets, IReadOnlyList<string> targetSheets) {
            SourceSheets = sourceSheets;
            TargetSheets = targetSheets;
        }

        /// <summary>
        /// Gets the imported source worksheet names.
        /// </summary>
        public IReadOnlyList<string> SourceSheets { get; }

        /// <summary>
        /// Gets the created target worksheet names.
        /// </summary>
        public IReadOnlyList<string> TargetSheets { get; }

        /// <summary>
        /// Gets the number of imported worksheets.
        /// </summary>
        public int SheetCount => TargetSheets.Count;
    }
}
