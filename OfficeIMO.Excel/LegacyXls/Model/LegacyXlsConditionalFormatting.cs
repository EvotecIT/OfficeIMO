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
            IReadOnlyList<string> ranges,
            int? priority = null,
            bool stopIfTrue = false,
            LegacyXlsDifferentialFormat? differentialFormat = null) {
            Type = type;
            Operator = @operator;
            Formula1 = formula1 ?? throw new ArgumentNullException(nameof(formula1));
            Formula2 = formula2;
            Ranges = ranges ?? throw new ArgumentNullException(nameof(ranges));
            Priority = priority;
            StopIfTrue = stopIfTrue;
            DifferentialFormat = differentialFormat;
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

        /// <summary>
        /// Gets the number of A1 ranges covered by this conditional formatting rule.
        /// </summary>
        public int RangeCount => Ranges.Count;

        /// <summary>
        /// Gets the evaluation priority supplied by a conditional-formatting extension record, when present.
        /// </summary>
        public int? Priority { get; private set; }

        /// <summary>
        /// Gets whether lower-priority rules should stop when this rule evaluates to true.
        /// </summary>
        public bool StopIfTrue { get; private set; }

        /// <summary>
        /// Gets the differential format associated with this rule, when a supported style extension was decoded.
        /// </summary>
        public LegacyXlsDifferentialFormat? DifferentialFormat { get; private set; }

        internal void ApplyExtension(int? priority, bool stopIfTrue, LegacyXlsDifferentialFormat? differentialFormat) {
            Priority = priority;
            StopIfTrue = stopIfTrue;
            if (differentialFormat != null) {
                DifferentialFormat = differentialFormat;
            }
        }
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
