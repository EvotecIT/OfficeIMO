namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls how workbook calculation metadata is handled when stale calculation chains are removed.
    /// </summary>
    public enum ExcelCalculationCleanupPolicy {
        /// <summary>
        /// Removes stale calculation-chain parts while preserving existing workbook calculation properties.
        /// </summary>
        PreserveExistingCalculationProperties,

        /// <summary>
        /// Removes stale calculation-chain parts and clears automatic full-recalculation flags from existing workbook calculation properties.
        /// </summary>
        ClearAutomaticFullCalculationOnOpen,

        /// <summary>
        /// Removes stale calculation-chain parts and explicitly requests full recalculation when the workbook is opened.
        /// </summary>
        RequestFullCalculationOnOpen
    }

    /// <summary>
    /// Controls how formula cells are treated before a workbook is saved.
    /// </summary>
    public sealed class ExcelCalculationOptions {
        private int _maximumDependencyDepth = 256;
        private readonly object _customFunctionLock = new object();
        private readonly Dictionary<string, ExcelCustomFormulaFunction> _customFunctions = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Gets or sets the maximum number of nested formula cells OfficeIMO evaluates in one dependency chain.
        /// </summary>
        /// <remarks>
        /// The default is 256. Formulas beyond the limit remain intact and are left for Excel-compatible applications to recalculate.
        /// </remarks>
        public int MaximumDependencyDepth {
            get => _maximumDependencyDepth;
            set {
                if (value < 1) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Formula dependency depth must be at least 1.");
                }

                _maximumDependencyDepth = value;
            }
        }

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

        /// <summary>
        /// Gets the registered custom function names in ordinal order.
        /// </summary>
        public IReadOnlyList<string> CustomFunctionNames {
            get {
                lock (_customFunctionLock) {
                    return _customFunctions.Keys.OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray();
                }
            }
        }

        /// <summary>
        /// Registers or replaces a custom formula function for this workbook.
        /// </summary>
        /// <remarks>
        /// Built-in OfficeIMO function names cannot be replaced. A custom function is an in-memory calculation callback and is not embedded in saved workbooks.
        /// </remarks>
        public void RegisterCustomFunction(string name, ExcelCustomFormulaFunction function) {
            string normalizedName = NormalizeCustomFunctionName(name);
            if (function == null) {
                throw new ArgumentNullException(nameof(function));
            }

            if (ExcelFormulaCapabilities.IsBuiltInFunction(normalizedName)) {
                throw new ArgumentException($"'{normalizedName}' is a built-in OfficeIMO formula function and cannot be replaced.", nameof(name));
            }

            lock (_customFunctionLock) {
                _customFunctions[normalizedName] = function;
            }
        }

        /// <summary>
        /// Removes a registered custom formula function.
        /// </summary>
        public bool RemoveCustomFunction(string name) {
            string normalizedName = NormalizeCustomFunctionName(name);
            lock (_customFunctionLock) {
                return _customFunctions.Remove(normalizedName);
            }
        }

        /// <summary>
        /// Removes all custom formula functions registered for this workbook.
        /// </summary>
        public void ClearCustomFunctions() {
            lock (_customFunctionLock) {
                _customFunctions.Clear();
            }
        }

        internal bool TryGetCustomFunction(string normalizedName, out ExcelCustomFormulaFunction? function) {
            lock (_customFunctionLock) {
                return _customFunctions.TryGetValue(normalizedName, out function);
            }
        }

        private static string NormalizeCustomFunctionName(string name) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentNullException(nameof(name));
            }

            string normalized = name.Trim().ToUpperInvariant();
            if (normalized.Length > 255) {
                throw new ArgumentException("Custom formula function names cannot exceed 255 characters.", nameof(name));
            }

            if (!IsCustomFunctionNameStart(normalized[0])) {
                throw new ArgumentException("Custom formula function names must start with a letter or underscore.", nameof(name));
            }

            for (int index = 1; index < normalized.Length; index++) {
                char character = normalized[index];
                if (!IsAsciiLetter(character) && !IsAsciiDigit(character) && character != '_' && character != '.') {
                    throw new ArgumentException("Custom formula function names may contain only letters, digits, underscores, and periods.", nameof(name));
                }
            }

            return normalized;
        }

        private static bool IsCustomFunctionNameStart(char character) {
            return IsAsciiLetter(character) || character == '_';
        }

        private static bool IsAsciiLetter(char character) {
            return character >= 'A' && character <= 'Z';
        }

        private static bool IsAsciiDigit(char character) {
            return character >= '0' && character <= '9';
        }
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
