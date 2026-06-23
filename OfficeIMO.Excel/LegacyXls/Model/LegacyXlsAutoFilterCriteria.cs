namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents parsed legacy AutoFilter criteria for one filtered column.
    /// </summary>
    public sealed class LegacyXlsAutoFilterCriteria {
        /// <summary>
        /// Creates parsed legacy AutoFilter criteria.
        /// </summary>
        public LegacyXlsAutoFilterCriteria(
            uint columnId,
            bool matchAll,
            IReadOnlyList<LegacyXlsAutoFilterCondition> conditions,
            LegacyXlsAutoFilterKind kind = LegacyXlsAutoFilterKind.Custom,
            ushort? top10Value = null,
            bool top10IsTop = true,
            bool top10IsPercent = false) {
            if (kind == LegacyXlsAutoFilterKind.Top10 && (!top10Value.HasValue || top10Value.Value < 1 || top10Value.Value > 500)) {
                throw new ArgumentOutOfRangeException(nameof(top10Value), "Top10 AutoFilter values must be between 1 and 500.");
            }

            ColumnId = columnId;
            MatchAll = matchAll;
            Conditions = conditions ?? throw new ArgumentNullException(nameof(conditions));
            Kind = kind;
            Top10Value = top10Value;
            Top10IsTop = top10IsTop;
            Top10IsPercent = top10IsPercent;
        }

        /// <summary>
        /// Creates parsed legacy Top/Bottom AutoFilter criteria.
        /// </summary>
        public static LegacyXlsAutoFilterCriteria CreateTop10(uint columnId, ushort value, bool isTop, bool isPercent) {
            return new LegacyXlsAutoFilterCriteria(
                columnId,
                matchAll: false,
                conditions: Array.Empty<LegacyXlsAutoFilterCondition>(),
                kind: LegacyXlsAutoFilterKind.Top10,
                top10Value: value,
                top10IsTop: isTop,
                top10IsPercent: isPercent);
        }

        /// <summary>
        /// Creates parsed legacy blank AutoFilter criteria.
        /// </summary>
        public static LegacyXlsAutoFilterCriteria CreateBlanks(uint columnId) {
            return new LegacyXlsAutoFilterCriteria(
                columnId,
                matchAll: false,
                conditions: new[] { new LegacyXlsAutoFilterCondition(LegacyXlsAutoFilterOperator.Equal, string.Empty) },
                kind: LegacyXlsAutoFilterKind.Blanks);
        }

        /// <summary>
        /// Creates parsed legacy nonblank AutoFilter criteria.
        /// </summary>
        public static LegacyXlsAutoFilterCriteria CreateNonBlanks(uint columnId) {
            return new LegacyXlsAutoFilterCriteria(
                columnId,
                matchAll: false,
                conditions: new[] { new LegacyXlsAutoFilterCondition(LegacyXlsAutoFilterOperator.NotEqual, string.Empty) },
                kind: LegacyXlsAutoFilterKind.NonBlanks);
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

        /// <summary>
        /// Gets the kind of AutoFilter criteria represented by this record.
        /// </summary>
        public LegacyXlsAutoFilterKind Kind { get; }

        /// <summary>
        /// Gets whether this criteria represents a Top/Bottom AutoFilter.
        /// </summary>
        public bool IsTop10 => Kind == LegacyXlsAutoFilterKind.Top10;

        /// <summary>
        /// Gets the Top/Bottom count or percentage value.
        /// </summary>
        public ushort? Top10Value { get; }

        /// <summary>
        /// Gets whether the Top/Bottom criteria keeps top values rather than bottom values.
        /// </summary>
        public bool Top10IsTop { get; }

        /// <summary>
        /// Gets whether the Top/Bottom criteria value is a percentage rather than an item count.
        /// </summary>
        public bool Top10IsPercent { get; }
    }

    /// <summary>
    /// Identifies the shape of parsed legacy AutoFilter criteria.
    /// </summary>
    public enum LegacyXlsAutoFilterKind {
        /// <summary>Comparison or equality-list criteria.</summary>
        Custom,
        /// <summary>Top or bottom item/percentage criteria.</summary>
        Top10,
        /// <summary>Blank-cell criteria.</summary>
        Blanks,
        /// <summary>Nonblank-cell criteria.</summary>
        NonBlanks
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
