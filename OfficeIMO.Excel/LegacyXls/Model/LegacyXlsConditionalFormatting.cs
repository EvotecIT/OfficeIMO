namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a conditional formatting rule parsed from a legacy BIFF worksheet.
    /// </summary>
    public sealed class LegacyXlsConditionalFormatting {
        /// <summary>
        /// Creates a parsed legacy XLS conditional formatting rule.
        /// </summary>
        public LegacyXlsConditionalFormatting(
            LegacyXlsConditionalFormattingType type,
            LegacyXlsConditionalFormattingOperator? @operator,
            string formula1,
            string? formula2,
            IReadOnlyList<string> ranges) {
            Type = type;
            Operator = @operator;
            Formula1 = formula1 ?? throw new ArgumentNullException(nameof(formula1));
            Formula2 = formula2;
            Ranges = ranges ?? throw new ArgumentNullException(nameof(ranges));
        }

        /// <summary>
        /// Gets the legacy conditional formatting rule type.
        /// </summary>
        public LegacyXlsConditionalFormattingType Type { get; }

        /// <summary>
        /// Gets the comparison operator for cell-is rules.
        /// </summary>
        public LegacyXlsConditionalFormattingOperator? Operator { get; }

        /// <summary>
        /// Gets the first formula, normalized for Open XML projection.
        /// </summary>
        public string Formula1 { get; }

        /// <summary>
        /// Gets the second formula, when required by the operator.
        /// </summary>
        public string? Formula2 { get; }

        /// <summary>
        /// Gets the A1 ranges covered by this conditional formatting rule.
        /// </summary>
        public IReadOnlyList<string> Ranges { get; }
    }

    /// <summary>
    /// Identifies a supported legacy conditional formatting rule type.
    /// </summary>
    public enum LegacyXlsConditionalFormattingType {
        /// <summary>
        /// Cell value comparison rule.
        /// </summary>
        CellIs,

        /// <summary>
        /// Formula expression rule.
        /// </summary>
        Formula
    }

    /// <summary>
    /// Identifies a legacy conditional formatting comparison operator.
    /// </summary>
    public enum LegacyXlsConditionalFormattingOperator {
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
