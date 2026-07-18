namespace OfficeIMO.Email.Store;

/// <summary>Sort direction for a Store table query.</summary>
public enum EmailStoreSortDirection {
    /// <summary>Lowest values first.</summary>
    Ascending,
    /// <summary>Highest values first.</summary>
    Descending
}

/// <summary>Placement of null values independent of sort direction.</summary>
public enum EmailStoreNullOrder {
    /// <summary>Null values precede non-null values.</summary>
    First,
    /// <summary>Null values follow non-null values.</summary>
    Last
}

/// <summary>One typed ordering clause in a Store table query.</summary>
public sealed class EmailStoreSort {
    /// <summary>Creates an ordering clause.</summary>
    public EmailStoreSort(EmailStoreField field, EmailStoreSortDirection direction = EmailStoreSortDirection.Ascending,
        EmailStoreNullOrder nullOrder = EmailStoreNullOrder.Last) {
        Field = field ?? throw new ArgumentNullException(nameof(field));
        Direction = direction;
        NullOrder = nullOrder;
    }

    /// <summary>Field being ordered.</summary>
    public EmailStoreField Field { get; }

    /// <summary>Sort direction.</summary>
    public EmailStoreSortDirection Direction { get; }

    /// <summary>Explicit null placement.</summary>
    public EmailStoreNullOrder NullOrder { get; }

    internal int Compare(EmailStoreQueryRow left, EmailStoreQueryRow right) =>
        CompareValues(Field.Read(left), Field.Read(right));

    internal int CompareValues(object? left, object? right) {
        if (left == null || right == null) {
            if (left == null && right == null) return 0;
            return left == null
                ? (NullOrder == EmailStoreNullOrder.First ? -1 : 1)
                : (NullOrder == EmailStoreNullOrder.First ? 1 : -1);
        }
        int result = Field.CompareNonNull(left, right);
        return Direction == EmailStoreSortDirection.Descending ? -result : result;
    }

    internal string Signature => string.Concat(Field.Key, ":", ((int)Direction).ToString(CultureInfo.InvariantCulture),
        ":", ((int)NullOrder).ToString(CultureInfo.InvariantCulture));
}
