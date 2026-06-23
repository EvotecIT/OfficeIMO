namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a data validation rule parsed from a legacy BIFF worksheet.
    /// </summary>
    public sealed class LegacyXlsDataValidation {
        /// <summary>
        /// Creates a parsed legacy XLS data validation rule.
        /// </summary>
        public LegacyXlsDataValidation(
            LegacyXlsDataValidationType type,
            LegacyXlsDataValidationOperator @operator,
            string formula1,
            string? formula2,
            bool allowBlank,
            bool showInputMessage,
            bool showErrorMessage,
            string? promptTitle,
            string? prompt,
            string? errorTitle,
            string? error,
            IReadOnlyList<string> ranges,
            IReadOnlyList<string>? listItems = null,
            string? listSourceRange = null,
            string? listSourceName = null,
            string? listSourceSheetName = null) {
            Type = type;
            Operator = @operator;
            Formula1 = formula1 ?? throw new ArgumentNullException(nameof(formula1));
            Formula2 = formula2;
            AllowBlank = allowBlank;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
            PromptTitle = promptTitle;
            Prompt = prompt;
            ErrorTitle = errorTitle;
            Error = error;
            Ranges = ranges ?? throw new ArgumentNullException(nameof(ranges));
            ListItems = listItems ?? Array.Empty<string>();
            ListSourceRange = listSourceRange;
            ListSourceName = listSourceName;
            ListSourceSheetName = listSourceSheetName;
        }

        /// <summary>
        /// Gets the legacy validation type.
        /// </summary>
        public LegacyXlsDataValidationType Type { get; }

        /// <summary>
        /// Gets the validation comparison operator.
        /// </summary>
        public LegacyXlsDataValidationOperator Operator { get; }

        /// <summary>
        /// Gets the first validation formula, normalized for Open XML projection.
        /// </summary>
        public string Formula1 { get; }

        /// <summary>
        /// Gets the second validation formula, when required by the operator.
        /// </summary>
        public string? Formula2 { get; }

        /// <summary>
        /// Gets whether blank values are accepted by the validation.
        /// </summary>
        public bool AllowBlank { get; }

        /// <summary>
        /// Gets whether Excel should show an input prompt for the validation.
        /// </summary>
        public bool ShowInputMessage { get; }

        /// <summary>
        /// Gets whether Excel should show an error prompt for invalid entries.
        /// </summary>
        public bool ShowErrorMessage { get; }

        /// <summary>
        /// Gets the input prompt title.
        /// </summary>
        public string? PromptTitle { get; }

        /// <summary>
        /// Gets the input prompt body.
        /// </summary>
        public string? Prompt { get; }

        /// <summary>
        /// Gets the error prompt title.
        /// </summary>
        public string? ErrorTitle { get; }

        /// <summary>
        /// Gets the error prompt body.
        /// </summary>
        public string? Error { get; }

        /// <summary>
        /// Gets the A1 ranges covered by this validation rule.
        /// </summary>
        public IReadOnlyList<string> Ranges { get; }

        /// <summary>
        /// Gets inline list items for list validation rules.
        /// </summary>
        public IReadOnlyList<string> ListItems { get; }

        /// <summary>
        /// Gets the same-sheet A1 source range for range-backed list validation rules.
        /// </summary>
        public string? ListSourceRange { get; }

        /// <summary>
        /// Gets the source sheet for sheet-qualified range-backed list validation rules.
        /// </summary>
        public string? ListSourceSheetName { get; }

        /// <summary>
        /// Gets the defined name source for named-range-backed list validation rules.
        /// </summary>
        public string? ListSourceName { get; }
    }

    /// <summary>
    /// Identifies a legacy data validation value type.
    /// </summary>
    public enum LegacyXlsDataValidationType {
        /// <summary>
        /// Whole-number validation.
        /// </summary>
        WholeNumber,

        /// <summary>
        /// Decimal-number validation.
        /// </summary>
        Decimal,

        /// <summary>
        /// Inline list validation.
        /// </summary>
        List,

        /// <summary>
        /// Date validation.
        /// </summary>
        Date,

        /// <summary>
        /// Time validation.
        /// </summary>
        Time,

        /// <summary>
        /// Text-length validation.
        /// </summary>
        TextLength,

        /// <summary>
        /// Custom formula validation.
        /// </summary>
        Custom
    }

    /// <summary>
    /// Identifies a legacy data validation comparison operator.
    /// </summary>
    public enum LegacyXlsDataValidationOperator {
        /// <summary>Between two bounds.</summary>
        Between,
        /// <summary>Not between two bounds.</summary>
        NotBetween,
        /// <summary>Equal to a value.</summary>
        Equal,
        /// <summary>Not equal to a value.</summary>
        NotEqual,
        /// <summary>Greater than a value.</summary>
        GreaterThan,
        /// <summary>Less than a value.</summary>
        LessThan,
        /// <summary>Greater than or equal to a value.</summary>
        GreaterThanOrEqual,
        /// <summary>Less than or equal to a value.</summary>
        LessThanOrEqual
    }
}
