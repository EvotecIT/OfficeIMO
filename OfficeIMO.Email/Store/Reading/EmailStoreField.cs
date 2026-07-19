namespace OfficeIMO.Email.Store;

/// <summary>Typed field that can participate in Store filters, sorting, and table projections.</summary>
public abstract class EmailStoreField {
    internal EmailStoreField(string key, string displayName, Type valueType) {
        if (string.IsNullOrWhiteSpace(key)) throw new ArgumentException("A field key cannot be empty.", nameof(key));
        Key = key;
        DisplayName = string.IsNullOrWhiteSpace(displayName) ? key : displayName;
        ValueType = valueType ?? throw new ArgumentNullException(nameof(valueType));
    }

    /// <summary>Stable vocabulary key used by query plans and continuation tokens.</summary>
    public string Key { get; }

    /// <summary>Human-readable column name.</summary>
    public string DisplayName { get; }

    /// <summary>Projected CLR value type.</summary>
    public Type ValueType { get; }

    internal abstract object? Read(EmailStoreQueryRow row);
    internal abstract int CompareNonNull(object left, object right);

    /// <summary>Creates an ascending sort with explicit null placement.</summary>
    public EmailStoreSort Ascending(EmailStoreNullOrder nullOrder = EmailStoreNullOrder.Last) =>
        new EmailStoreSort(this, EmailStoreSortDirection.Ascending, nullOrder);

    /// <summary>Creates a descending sort with explicit null placement.</summary>
    public EmailStoreSort Descending(EmailStoreNullOrder nullOrder = EmailStoreNullOrder.Last) =>
        new EmailStoreSort(this, EmailStoreSortDirection.Descending, nullOrder);

    /// <inheritdoc />
    public override string ToString() => Key;
}

/// <summary>Typed Store field with comparison-filter helpers.</summary>
/// <typeparam name="T">Projected field type.</typeparam>
public class EmailStoreField<T> : EmailStoreField {
    private readonly Func<EmailStoreQueryRow, T> _selector;
    private readonly IComparer<T> _comparer;

    internal EmailStoreField(string key, string displayName, Func<EmailStoreQueryRow, T> selector,
        IComparer<T>? comparer = null)
        : base(key, displayName, typeof(T)) {
        _selector = selector ?? throw new ArgumentNullException(nameof(selector));
        _comparer = comparer ?? Comparer<T>.Default;
    }

    /// <summary>Tests ordinal/typed equality according to this field's vocabulary comparer.</summary>
    public EmailStoreFilter EqualTo(T value) => EmailStoreFilter.Comparison(this, EmailStoreComparisonOperator.Equal, value);

    /// <summary>Tests inequality according to this field's vocabulary comparer.</summary>
    public EmailStoreFilter NotEqualTo(T value) => EmailStoreFilter.Comparison(this, EmailStoreComparisonOperator.NotEqual, value);

    /// <summary>Tests whether the field sorts after the supplied value.</summary>
    public EmailStoreFilter GreaterThan(T value) => EmailStoreFilter.Comparison(this, EmailStoreComparisonOperator.GreaterThan, value);

    /// <summary>Tests whether the field sorts at or after the supplied value.</summary>
    public EmailStoreFilter GreaterThanOrEqualTo(T value) => EmailStoreFilter.Comparison(this, EmailStoreComparisonOperator.GreaterThanOrEqual, value);

    /// <summary>Tests whether the field sorts before the supplied value.</summary>
    public EmailStoreFilter LessThan(T value) => EmailStoreFilter.Comparison(this, EmailStoreComparisonOperator.LessThan, value);

    /// <summary>Tests whether the field sorts at or before the supplied value.</summary>
    public EmailStoreFilter LessThanOrEqualTo(T value) => EmailStoreFilter.Comparison(this, EmailStoreComparisonOperator.LessThanOrEqual, value);

    /// <summary>Tests whether the field is null.</summary>
    public EmailStoreFilter IsNull() => EmailStoreFilter.Comparison(this, EmailStoreComparisonOperator.IsNull, null);

    /// <summary>Tests whether the field is non-null.</summary>
    public EmailStoreFilter IsNotNull() => EmailStoreFilter.Comparison(this, EmailStoreComparisonOperator.IsNotNull, null);

    /// <summary>Tests membership using the field's vocabulary comparer.</summary>
    public EmailStoreFilter In(params T[] values) => EmailStoreFilter.In(this, values ?? throw new ArgumentNullException(nameof(values)));

    internal override object? Read(EmailStoreQueryRow row) => _selector(row);

    internal override int CompareNonNull(object left, object right) => _comparer.Compare((T)left, (T)right);
}

/// <summary>String Store field with explicit case-insensitive matching helpers.</summary>
public sealed class EmailStoreStringField : EmailStoreField<string?> {
    internal EmailStoreStringField(string key, string displayName, Func<EmailStoreQueryRow, string?> selector)
        : base(key, displayName, selector, StringComparer.OrdinalIgnoreCase) {
    }

    /// <summary>Tests for a case-insensitive substring.</summary>
    public EmailStoreFilter Contains(string value) => EmailStoreFilter.String(this, EmailStoreStringOperator.Contains, value);

    /// <summary>Tests for a case-insensitive prefix.</summary>
    public EmailStoreFilter StartsWith(string value) => EmailStoreFilter.String(this, EmailStoreStringOperator.StartsWith, value);

    /// <summary>Tests for a case-insensitive suffix.</summary>
    public EmailStoreFilter EndsWith(string value) => EmailStoreFilter.String(this, EmailStoreStringOperator.EndsWith, value);
}

internal sealed class EmailStoreQueryRow {
    internal EmailStoreQueryRow(EmailStoreItemReference reference, EmailStoreItemSummary summary) {
        Reference = reference;
        Summary = summary;
    }

    internal EmailStoreItemReference Reference { get; }
    internal EmailStoreItemSummary Summary { get; }
}
