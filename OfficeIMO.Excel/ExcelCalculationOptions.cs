namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls how formula cells are treated before a workbook is saved.
    /// </summary>
    public sealed class ExcelCalculationOptions {
        /// <summary>
        /// When true, OfficeIMO evaluates supported formulas and writes cached values before saving.
        /// Unsupported formulas are left intact and can still be recalculated by Excel.
        /// </summary>
        public bool EvaluateFormulasBeforeSave { get; set; }

        /// <summary>
        /// When true, formula cells are marked dirty so Excel-compatible applications recalculate them on open.
        /// </summary>
        public bool MarkFormulasDirtyBeforeSave { get; set; }

        /// <summary>
        /// When true, workbook calculation properties request a full recalculation when the file opens.
        /// </summary>
        public bool ForceFullCalculationOnOpen { get; set; }

        /// <summary>
        /// When true, cached formula results are removed before saving. Ignored when <see cref="EvaluateFormulasBeforeSave"/> is true.
        /// </summary>
        public bool ClearCachedFormulaResultsBeforeSave { get; set; }
    }

    /// <summary>
    /// Options used when protecting the workbook structure.
    /// </summary>
    public sealed class ExcelWorkbookProtectionOptions {
        /// <summary>
        /// Protects the workbook structure, which prevents sheet add/delete/move/rename operations in Excel UI.
        /// </summary>
        public bool ProtectStructure { get; set; } = true;

        /// <summary>
        /// Protects workbook windows where supported by the consuming application.
        /// </summary>
        public bool ProtectWindows { get; set; }

        /// <summary>
        /// Optional workbook protection password. This is Excel UI protection, not package encryption.
        /// </summary>
        public string? Password { get; set; }

        /// <summary>
        /// Optional precomputed legacy workbook protection hash. When set, this value is written as-is.
        /// </summary>
        public string? LegacyPasswordHash { get; set; }
    }

    /// <summary>
    /// Selects which cell parts are cleared by <see cref="ExcelRange.Clear"/>.
    /// </summary>
    [System.Flags]
    public enum ExcelClearOptions {
        /// <summary>Do not clear any cell data or metadata.</summary>
        None = 0,
        /// <summary>Clear literal cell values.</summary>
        Values = 1,
        /// <summary>Clear cell formulas.</summary>
        Formulas = 2,
        /// <summary>Clear cell style indexes.</summary>
        Styles = 4,
        /// <summary>Clear comments in the selected cells when supported.</summary>
        Comments = 8,
        /// <summary>Clear hyperlinks that overlap the range.</summary>
        Hyperlinks = 16,
        /// <summary>Clear data validation rules that overlap the range.</summary>
        DataValidations = 32,
        /// <summary>Clear conditional formatting rules that overlap the range.</summary>
        ConditionalFormatting = 64,
        /// <summary>Clear merged-cell definitions that overlap the range.</summary>
        Merges = 128,
        /// <summary>Clear sparklines whose target cells overlap the range.</summary>
        Sparklines = 256,
        /// <summary>Clear all supported cell data and range metadata.</summary>
        All = Values | Formulas | Styles | Comments | Hyperlinks | DataValidations | ConditionalFormatting | Merges | Sparklines
    }

    /// <summary>
    /// Kind of value stored in an <see cref="ExcelCellData"/>.
    /// </summary>
    public enum ExcelCellDataKind {
        /// <summary>The cell has no value.</summary>
        Blank,
        /// <summary>The cell contains a Boolean value.</summary>
        Boolean,
        /// <summary>The cell contains a numeric value.</summary>
        Number,
        /// <summary>The cell contains text.</summary>
        Text,
        /// <summary>The cell contains an error value.</summary>
        Error,
        /// <summary>The cell contains a formula.</summary>
        Formula
    }

    /// <summary>
    /// A typed snapshot of a worksheet cell value.
    /// </summary>
    public sealed class ExcelCellData {
        /// <summary>
        /// Creates a snapshot for a worksheet cell value.
        /// </summary>
        public ExcelCellData(ExcelCellDataKind kind, object? value, string? formula = null, string? cachedText = null) {
            Kind = kind;
            Value = value;
            Formula = formula;
            CachedText = cachedText;
        }

        /// <summary>
        /// Gets the value kind.
        /// </summary>
        public ExcelCellDataKind Kind { get; }

        /// <summary>
        /// Gets the typed value when one is available.
        /// </summary>
        public object? Value { get; }

        /// <summary>
        /// Gets the formula text for formula cells.
        /// </summary>
        public string? Formula { get; }

        /// <summary>
        /// Gets the cached text from the cell value.
        /// </summary>
        public string? CachedText { get; }

        /// <summary>
        /// Gets a value indicating whether the cell is blank.
        /// </summary>
        public bool IsBlank => Kind == ExcelCellDataKind.Blank;

        /// <summary>
        /// Gets a value indicating whether the cell contains a formula.
        /// </summary>
        public bool HasFormula => Kind == ExcelCellDataKind.Formula;
    }
}
