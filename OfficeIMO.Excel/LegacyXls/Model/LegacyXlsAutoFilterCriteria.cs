namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents parsed legacy AutoFilter criteria for one filtered column.
    /// </summary>
    public sealed class LegacyXlsAutoFilterCriteria {
        /// <summary>
        /// Creates parsed legacy AutoFilter criteria.
        /// </summary>
        public LegacyXlsAutoFilterCriteria(uint columnId, bool matchAll, IReadOnlyList<LegacyXlsAutoFilterCondition> conditions) {
            ColumnId = columnId;
            MatchAll = matchAll;
            Conditions = conditions ?? throw new ArgumentNullException(nameof(conditions));
        }

        /// <summary>
        /// Gets the zero-based column index within the AutoFilter range.
        /// </summary>
        public uint ColumnId { get; }

        /// <summary>
        /// Gets whether multiple conditions must all match.
        /// </summary>
        public bool MatchAll { get; }

        /// <summary>
        /// Gets the parsed filter conditions.
        /// </summary>
        public IReadOnlyList<LegacyXlsAutoFilterCondition> Conditions { get; }
    }

    /// <summary>
    /// Represents one parsed legacy AutoFilter condition.
    /// </summary>
    public sealed class LegacyXlsAutoFilterCondition {
        /// <summary>
        /// Creates a parsed legacy AutoFilter condition.
        /// </summary>
        public LegacyXlsAutoFilterCondition(LegacyXlsAutoFilterOperator @operator, string value) {
            Operator = @operator;
            Value = value ?? throw new ArgumentNullException(nameof(value));
        }

        /// <summary>
        /// Gets the comparison operator.
        /// </summary>
        public LegacyXlsAutoFilterOperator Operator { get; }

        /// <summary>
        /// Gets the comparison value normalized for Open XML projection.
        /// </summary>
        public string Value { get; }
    }

    /// <summary>
    /// Identifies a legacy AutoFilter comparison operator.
    /// </summary>
    public enum LegacyXlsAutoFilterOperator {
        /// <summary>Less than.</summary>
        LessThan,
        /// <summary>Equal to.</summary>
        Equal,
        /// <summary>Less than or equal to.</summary>
        LessThanOrEqual,
        /// <summary>Greater than.</summary>
        GreaterThan,
        /// <summary>Not equal to.</summary>
        NotEqual,
        /// <summary>Greater than or equal to.</summary>
        GreaterThanOrEqual
    }
}
