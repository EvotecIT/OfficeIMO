namespace OfficeIMO.Email.Store;

/// <summary>Node kind in an immutable Store query expression tree.</summary>
public enum EmailStoreFilterKind {
    /// <summary>Matches every row.</summary>
    All,
    /// <summary>Typed scalar comparison.</summary>
    Comparison,
    /// <summary>Case-insensitive string operation.</summary>
    String,
    /// <summary>Set membership.</summary>
    In,
    /// <summary>Every child must match.</summary>
    And,
    /// <summary>At least one child must match.</summary>
    Or,
    /// <summary>Negates one child.</summary>
    Not
}

/// <summary>Typed scalar comparison operation.</summary>
public enum EmailStoreComparisonOperator {
    /// <summary>Values compare equal.</summary>
    Equal,
    /// <summary>Values do not compare equal.</summary>
    NotEqual,
    /// <summary>Field value sorts after the operand.</summary>
    GreaterThan,
    /// <summary>Field value sorts at or after the operand.</summary>
    GreaterThanOrEqual,
    /// <summary>Field value sorts before the operand.</summary>
    LessThan,
    /// <summary>Field value sorts at or before the operand.</summary>
    LessThanOrEqual,
    /// <summary>Field value is null.</summary>
    IsNull,
    /// <summary>Field value is non-null.</summary>
    IsNotNull
}

/// <summary>Case-insensitive string-match operation.</summary>
public enum EmailStoreStringOperator {
    /// <summary>Substring match.</summary>
    Contains,
    /// <summary>Prefix match.</summary>
    StartsWith,
    /// <summary>Suffix match.</summary>
    EndsWith
}

/// <summary>Immutable, composable Store query expression.</summary>
public abstract class EmailStoreFilter {
    private static readonly IReadOnlyList<EmailStoreFilter> NoChildren = Array.Empty<EmailStoreFilter>();
    private static readonly IReadOnlyList<object?> NoOperands = Array.Empty<object?>();

    /// <summary>Expression node kind.</summary>
    public abstract EmailStoreFilterKind Kind { get; }

    /// <summary>Canonical field for a leaf node, or null for a logical node.</summary>
    public virtual EmailStoreField? Field => null;

    /// <summary>Logical child nodes.</summary>
    public virtual IReadOnlyList<EmailStoreFilter> Children => NoChildren;

    /// <summary>Scalar operands for a comparison, string, or membership node.</summary>
    public virtual IReadOnlyList<object?> Operands => NoOperands;

    /// <summary>Comparison operation for a comparison node.</summary>
    public virtual EmailStoreComparisonOperator? ComparisonOperator => null;

    /// <summary>String operation for a string node.</summary>
    public virtual EmailStoreStringOperator? StringOperator => null;

    /// <summary>A filter that matches every row.</summary>
    public static EmailStoreFilter All { get; } = new AllFilter();

    /// <summary>Combines filters with logical AND.</summary>
    public static EmailStoreFilter And(params EmailStoreFilter[] filters) => Combine(EmailStoreFilterKind.And, filters);

    /// <summary>Combines filters with logical OR.</summary>
    public static EmailStoreFilter Or(params EmailStoreFilter[] filters) => Combine(EmailStoreFilterKind.Or, filters);

    /// <summary>Negates a filter.</summary>
    public static EmailStoreFilter Not(EmailStoreFilter filter) =>
        new CompositeFilter(EmailStoreFilterKind.Not, new[] { filter ?? throw new ArgumentNullException(nameof(filter)) });

    /// <summary>Combines two filters with logical AND.</summary>
    public static EmailStoreFilter operator &(EmailStoreFilter left, EmailStoreFilter right) => And(left, right);

    /// <summary>Combines two filters with logical OR.</summary>
    public static EmailStoreFilter operator |(EmailStoreFilter left, EmailStoreFilter right) => Or(left, right);

    /// <summary>Negates a filter.</summary>
    public static EmailStoreFilter operator !(EmailStoreFilter filter) => Not(filter);

    internal abstract bool Evaluate(EmailStoreQueryRow row);
    internal abstract string Signature { get; }

    internal static EmailStoreFilter Comparison(EmailStoreField field, EmailStoreComparisonOperator operation, object? value) =>
        new ComparisonFilter(field, operation, value);

    internal static EmailStoreFilter String(EmailStoreStringField field, EmailStoreStringOperator operation, string value) =>
        new StringFilter(field, operation, value);

    internal static EmailStoreFilter In<T>(EmailStoreField<T> field, IReadOnlyList<T> values) =>
        new InFilter(field, values.Cast<object?>().ToArray());

    private static EmailStoreFilter Combine(EmailStoreFilterKind kind, EmailStoreFilter[] filters) {
        if (filters == null) throw new ArgumentNullException(nameof(filters));
        if (filters.Length == 0) return All;
        if (filters.Any(filter => filter == null)) throw new ArgumentException("A filter collection cannot contain null.", nameof(filters));
        return filters.Length == 1 ? filters[0] : new CompositeFilter(kind, filters.ToArray());
    }

    private sealed class AllFilter : EmailStoreFilter {
        public override EmailStoreFilterKind Kind => EmailStoreFilterKind.All;
        internal override bool Evaluate(EmailStoreQueryRow row) => true;
        internal override string Signature => "all";
    }

    private sealed class ComparisonFilter : EmailStoreFilter {
        private readonly EmailStoreField _field;
        private readonly EmailStoreComparisonOperator _operation;
        private readonly object? _value;

        internal ComparisonFilter(EmailStoreField field, EmailStoreComparisonOperator operation, object? value) {
            _field = field ?? throw new ArgumentNullException(nameof(field));
            _operation = operation;
            _value = value;
        }

        public override EmailStoreFilterKind Kind => EmailStoreFilterKind.Comparison;
        public override EmailStoreField Field => _field;
        public override IReadOnlyList<object?> Operands => new[] { _value };
        public override EmailStoreComparisonOperator? ComparisonOperator => _operation;

        internal override bool Evaluate(EmailStoreQueryRow row) {
            object? actual = _field.Read(row);
            if (_operation == EmailStoreComparisonOperator.IsNull) return actual == null;
            if (_operation == EmailStoreComparisonOperator.IsNotNull) return actual != null;
            if (actual == null || _value == null) {
                bool equal = actual == null && _value == null;
                return _operation == EmailStoreComparisonOperator.Equal ? equal :
                    _operation == EmailStoreComparisonOperator.NotEqual && !equal;
            }
            int comparison = _field.CompareNonNull(actual, _value);
            switch (_operation) {
                case EmailStoreComparisonOperator.Equal: return comparison == 0;
                case EmailStoreComparisonOperator.NotEqual: return comparison != 0;
                case EmailStoreComparisonOperator.GreaterThan: return comparison > 0;
                case EmailStoreComparisonOperator.GreaterThanOrEqual: return comparison >= 0;
                case EmailStoreComparisonOperator.LessThan: return comparison < 0;
                case EmailStoreComparisonOperator.LessThanOrEqual: return comparison <= 0;
                default: throw new InvalidOperationException("Unsupported comparison operation.");
            }
        }

        internal override string Signature => string.Concat("cmp(", _field.Key, ",", ((int)_operation).ToString(CultureInfo.InvariantCulture), ",", EmailStoreScalarCodec.Signature(_value), ")");
    }

    private sealed class StringFilter : EmailStoreFilter {
        private readonly EmailStoreStringField _field;
        private readonly EmailStoreStringOperator _operation;
        private readonly string _value;

        internal StringFilter(EmailStoreStringField field, EmailStoreStringOperator operation, string value) {
            _field = field ?? throw new ArgumentNullException(nameof(field));
            if (string.IsNullOrEmpty(value)) throw new ArgumentException("A string filter value cannot be empty.", nameof(value));
            _operation = operation;
            _value = value;
        }

        public override EmailStoreFilterKind Kind => EmailStoreFilterKind.String;
        public override EmailStoreField Field => _field;
        public override IReadOnlyList<object?> Operands => new object?[] { _value };
        public override EmailStoreStringOperator? StringOperator => _operation;

        internal override bool Evaluate(EmailStoreQueryRow row) {
            string? actual = (string?)_field.Read(row);
            if (actual == null) return false;
            switch (_operation) {
                case EmailStoreStringOperator.Contains: return actual.IndexOf(_value, StringComparison.OrdinalIgnoreCase) >= 0;
                case EmailStoreStringOperator.StartsWith: return actual.StartsWith(_value, StringComparison.OrdinalIgnoreCase);
                case EmailStoreStringOperator.EndsWith: return actual.EndsWith(_value, StringComparison.OrdinalIgnoreCase);
                default: throw new InvalidOperationException("Unsupported string operation.");
            }
        }

        internal override string Signature => string.Concat("str(", _field.Key, ",", ((int)_operation).ToString(CultureInfo.InvariantCulture), ",", EmailStoreScalarCodec.Signature(_value), ")");
    }

    private sealed class InFilter : EmailStoreFilter {
        private readonly EmailStoreField _field;
        private readonly IReadOnlyList<object?> _values;

        internal InFilter(EmailStoreField field, IReadOnlyList<object?> values) {
            _field = field ?? throw new ArgumentNullException(nameof(field));
            _values = values ?? throw new ArgumentNullException(nameof(values));
        }

        public override EmailStoreFilterKind Kind => EmailStoreFilterKind.In;
        public override EmailStoreField Field => _field;
        public override IReadOnlyList<object?> Operands => _values;

        internal override bool Evaluate(EmailStoreQueryRow row) {
            object? actual = _field.Read(row);
            foreach (object? value in _values) {
                if (actual == null || value == null) {
                    if (actual == null && value == null) return true;
                } else if (_field.CompareNonNull(actual, value) == 0) {
                    return true;
                }
            }
            return false;
        }

        internal override string Signature => string.Concat("in(", _field.Key, ",",
            string.Join(",", _values.Select(EmailStoreScalarCodec.Signature)), ")");
    }

    private sealed class CompositeFilter : EmailStoreFilter {
        private readonly EmailStoreFilterKind _kind;
        private readonly IReadOnlyList<EmailStoreFilter> _children;

        internal CompositeFilter(EmailStoreFilterKind kind, IReadOnlyList<EmailStoreFilter> children) {
            _kind = kind;
            _children = children;
        }

        public override EmailStoreFilterKind Kind => _kind;
        public override IReadOnlyList<EmailStoreFilter> Children => _children;

        internal override bool Evaluate(EmailStoreQueryRow row) {
            if (_kind == EmailStoreFilterKind.Not) return !_children[0].Evaluate(row);
            if (_kind == EmailStoreFilterKind.And) return _children.All(child => child.Evaluate(row));
            return _children.Any(child => child.Evaluate(row));
        }

        internal override string Signature => string.Concat(_kind.ToString().ToLowerInvariant(), "(",
            string.Join(",", _children.Select(child => child.Signature)), ")");
    }
}
