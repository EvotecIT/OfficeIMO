using System.Collections.Generic;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for importing worksheets from one workbook into another.
    /// </summary>
    public sealed class ExcelWorkbookMergeOptions {
        /// <summary>
        /// Gets or sets source worksheet names to import. When empty, all source worksheets are imported.
        /// </summary>
        public IReadOnlyList<string>? SheetNames { get; set; }

        /// <summary>
        /// Gets or sets an optional prefix added to every imported worksheet name.
        /// </summary>
        public string? SheetNamePrefix { get; set; }

        /// <summary>
        /// Gets or sets sheet name validation behavior for imported worksheets.
        /// </summary>
        public SheetNameValidationMode SheetNameValidationMode { get; set; } = SheetNameValidationMode.Sanitize;
    }
}
