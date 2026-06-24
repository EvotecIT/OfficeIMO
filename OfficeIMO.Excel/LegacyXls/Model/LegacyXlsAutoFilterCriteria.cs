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
            JoinOperator = ResolveJoinOperator(Conditions, MatchAll);
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
                conditions: new[] { new LegacyXlsAutoFilterCondition(LegacyXlsAutoFilterOperator.Equal, string.Empty, LegacyXlsAutoFilterValueKind.Blank) },
                kind: LegacyXlsAutoFilterKind.Blanks);
        }

        /// <summary>
        /// Creates parsed legacy nonblank AutoFilter criteria.
        /// </summary>
        public static LegacyXlsAutoFilterCriteria CreateNonBlanks(uint columnId) {
            return new LegacyXlsAutoFilterCriteria(
                columnId,
                matchAll: false,
                conditions: new[] { new LegacyXlsAutoFilterCondition(LegacyXlsAutoFilterOperator.NotEqual, string.Empty, LegacyXlsAutoFilterValueKind.NonBlank) },
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
        /// Gets the logical join used when this criteria has comparison conditions.
        /// </summary>
        public LegacyXlsAutoFilterJoinOperator JoinOperator { get; }

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

        private static LegacyXlsAutoFilterJoinOperator ResolveJoinOperator(IReadOnlyList<LegacyXlsAutoFilterCondition> conditions, bool matchAll) {
            if (conditions.Count == 0) {
                return LegacyXlsAutoFilterJoinOperator.None;
            }

            if (conditions.Count == 1) {
                return LegacyXlsAutoFilterJoinOperator.Single;
            }

            return matchAll ? LegacyXlsAutoFilterJoinOperator.And : LegacyXlsAutoFilterJoinOperator.Or;
        }
    }

    /// <summary>
    /// Identifies how multiple AutoFilter comparison conditions are joined.
    /// </summary>
    public enum LegacyXlsAutoFilterJoinOperator {
        /// <summary>The criteria shape does not use comparison conditions.</summary>
        None,
        /// <summary>The criteria has one comparison condition.</summary>
        Single,
        /// <summary>All comparison conditions must match.</summary>
        And,
        /// <summary>Any comparison condition may match.</summary>
        Or
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
        public LegacyXlsAutoFilterCondition(
            LegacyXlsAutoFilterOperator @operator,
            string value,
            LegacyXlsAutoFilterValueKind valueKind = LegacyXlsAutoFilterValueKind.Unknown) {
            Operator = @operator;
            Value = value ?? throw new ArgumentNullException(nameof(value));
            ValueKind = valueKind;
            TextPatternKind = valueKind == LegacyXlsAutoFilterValueKind.Text
                ? GetTextPatternKind(value)
                : LegacyXlsAutoFilterTextPatternKind.None;
        }

        /// <summary>
        /// Gets the comparison operator.
        /// </summary>
        public LegacyXlsAutoFilterOperator Operator { get; }

        /// <summary>
        /// Gets the comparison value normalized for Open XML projection.
        /// </summary>
        public string Value { get; }

        /// <summary>
        /// Gets the BIFF operand kind used to store the comparison value.
        /// </summary>
        public LegacyXlsAutoFilterValueKind ValueKind { get; }

        /// <summary>
        /// Gets the wildcard-pattern shape for text operands.
        /// </summary>
        public LegacyXlsAutoFilterTextPatternKind TextPatternKind { get; }

        /// <summary>
        /// Gets whether the text operand contains an unescaped BIFF wildcard.
        /// </summary>
        public bool HasTextWildcardPattern => TextPatternKind != LegacyXlsAutoFilterTextPatternKind.None
            && TextPatternKind != LegacyXlsAutoFilterTextPatternKind.ExactText;

        private static LegacyXlsAutoFilterTextPatternKind GetTextPatternKind(string value) {
            int firstWildcard = IndexOfUnescapedWildcard(value);
            if (firstWildcard < 0) {
                return LegacyXlsAutoFilterTextPatternKind.ExactText;
            }

            bool startsWithWildcard = IsUnescapedWildcardAt(value, 0, '*');
            bool endsWithWildcard = IsUnescapedWildcardAt(value, value.Length - 1, '*');
            bool hasOnlyBoundaryAsterisks = HasOnlyBoundaryAsteriskWildcards(value, startsWithWildcard, endsWithWildcard);
            if (!hasOnlyBoundaryAsterisks) {
                return LegacyXlsAutoFilterTextPatternKind.WildcardExpression;
            }

            if (startsWithWildcard && endsWithWildcard && value.Length > 1) {
                return LegacyXlsAutoFilterTextPatternKind.Contains;
            }

            if (startsWithWildcard) {
                return LegacyXlsAutoFilterTextPatternKind.EndsWith;
            }

            return endsWithWildcard
                ? LegacyXlsAutoFilterTextPatternKind.BeginsWith
                : LegacyXlsAutoFilterTextPatternKind.WildcardExpression;
        }

        private static int IndexOfUnescapedWildcard(string value) {
            for (int i = 0; i < value.Length; i++) {
                if ((value[i] == '*' || value[i] == '?') && !IsEscaped(value, i)) {
                    return i;
                }
            }

            return -1;
        }

        private static bool HasOnlyBoundaryAsteriskWildcards(string value, bool startsWithWildcard, bool endsWithWildcard) {
            for (int i = 0; i < value.Length; i++) {
                if ((value[i] == '*' || value[i] == '?') && !IsEscaped(value, i)) {
                    bool boundaryAsterisk = value[i] == '*'
                        && ((startsWithWildcard && i == 0) || (endsWithWildcard && i == value.Length - 1));
                    if (!boundaryAsterisk) {
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool IsUnescapedWildcardAt(string value, int index, char wildcard) {
            return index >= 0
                && index < value.Length
                && value[index] == wildcard
                && !IsEscaped(value, index);
        }

        private static bool IsEscaped(string value, int index) {
            int escapeCount = 0;
            for (int i = index - 1; i >= 0 && value[i] == '~'; i--) {
                escapeCount++;
            }

            return (escapeCount % 2) != 0;
        }
    }

    /// <summary>
    /// Identifies the BIFF operand kind used by a legacy AutoFilter condition.
    /// </summary>
    public enum LegacyXlsAutoFilterValueKind {
        /// <summary>The operand kind was not specified by the caller.</summary>
        Unknown,
        /// <summary>The operand was stored as an RK number.</summary>
        RkNumber,
        /// <summary>The operand was stored as an IEEE 754 number.</summary>
        Number,
        /// <summary>The operand was stored as text.</summary>
        Text,
        /// <summary>The operand was stored as a Boolean or error value.</summary>
        BooleanOrError,
        /// <summary>The operand was the blank-cell sentinel.</summary>
        Blank,
        /// <summary>The operand was the nonblank-cell sentinel.</summary>
        NonBlank
    }

    /// <summary>
    /// Identifies the text-pattern shape represented by a legacy AutoFilter text operand.
    /// </summary>
    public enum LegacyXlsAutoFilterTextPatternKind {
        /// <summary>The condition is not a text operand.</summary>
        None,
        /// <summary>The text operand does not contain BIFF wildcard characters.</summary>
        ExactText,
        /// <summary>The text operand ends with an unescaped asterisk wildcard.</summary>
        BeginsWith,
        /// <summary>The text operand starts with an unescaped asterisk wildcard.</summary>
        EndsWith,
        /// <summary>The text operand starts and ends with unescaped asterisk wildcards.</summary>
        Contains,
        /// <summary>The text operand contains a wildcard expression that is not one of the simple prefix/suffix shapes.</summary>
        WildcardExpression
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
