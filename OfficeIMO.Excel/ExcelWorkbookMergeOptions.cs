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

        /// <summary>
        /// Gets or sets the worksheet copy strategy used for imported sheets.
        /// </summary>
        public ExcelWorksheetCopyMode CopyMode { get; set; } = ExcelWorksheetCopyMode.Values;

        /// <summary>
        /// Gets or sets whether package-mode imports may preserve external-workbook references.
        /// External links are rejected unless callers opt in explicitly.
        /// </summary>
        public bool CopyExternalWorkbookReferences { get; set; }

        /// <summary>
        /// Gets or sets the maximum number of referenced defined names copied in package mode.
        /// </summary>
        public int MaxDefinedNames { get; set; } = 4096;

        /// <summary>
        /// Gets or sets the maximum aggregate formula characters across copied defined names.
        /// </summary>
        public int MaxDefinedNameCharacters { get; set; } = 1_000_000;
    }
}
