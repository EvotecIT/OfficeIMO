namespace OfficeIMO.Email.Store;

/// <summary>One projected row from a Store table query.</summary>
public sealed class EmailStoreTableRow {
    private readonly IReadOnlyList<EmailStoreField> _fields;
    private readonly IReadOnlyList<object?> _values;
    private readonly Dictionary<string, int> _indexes;

    internal EmailStoreTableRow(EmailStoreQueryRow row, EmailStoreProjection projection) {
        Reference = row.Reference;
        Summary = row.Summary;
        _fields = projection.Fields;
        var values = new object?[projection.Fields.Count];
        _indexes = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int index = 0; index < values.Length; index++) {
            EmailStoreField field = projection.Fields[index];
            values[index] = field.Read(row);
            _indexes.Add(field.Key, index);
        }
        _values = Array.AsReadOnly(values);
    }

    /// <summary>Stable reference that can be passed to selective read APIs.</summary>
    public EmailStoreItemReference Reference { get; }

    /// <summary>Lightweight summary used to evaluate this row.</summary>
    public EmailStoreItemSummary Summary { get; }

    /// <summary>Projected fields in column order.</summary>
    public IReadOnlyList<EmailStoreField> Fields => _fields;

    /// <summary>Projected values in column order.</summary>
    public IReadOnlyList<object?> Values => _values;

    /// <summary>Gets a projected value by canonical field.</summary>
    public object? this[EmailStoreField field] {
        get {
            if (field == null) throw new ArgumentNullException(nameof(field));
            if (!_indexes.TryGetValue(field.Key, out int index)) {
                throw new KeyNotFoundException(string.Concat("Field '", field.Key, "' is not in this projection."));
            }
            return _values[index];
        }
    }

    /// <summary>Gets a projected value with compile-time typing.</summary>
    public T Get<T>(EmailStoreField<T> field) {
        object? value = this[field];
        if (value == null) return default!;
        return (T)value;
    }
}

/// <summary>Bounded, deterministically ordered page from a Store table query.</summary>
public sealed class EmailStoreTablePage {
    internal EmailStoreTablePage(IReadOnlyList<EmailStoreTableRow> rows,
        EmailStoreContinuationToken? nextToken, int itemsScanned, int matchesInScan,
        bool scanLimitReached, EmailStoreQueryPlan plan) {
        Rows = rows;
        NextToken = nextToken;
        ItemsScanned = itemsScanned;
        MatchesInScan = matchesInScan;
        ScanLimitReached = scanLimitReached;
        Plan = plan;
    }

    /// <summary>Projected rows.</summary>
    public IReadOnlyList<EmailStoreTableRow> Rows { get; }

    /// <summary>Continuation for the next page, or null when this scanned result set is exhausted.</summary>
    public EmailStoreContinuationToken? NextToken { get; }

    /// <summary>Lightweight references evaluated during this execution.</summary>
    public int ItemsScanned { get; }

    /// <summary>Total matching rows within the bounded scan, including rows before this continuation.</summary>
    public int MatchesInScan { get; }

    /// <summary>True when more source references existed beyond <see cref="EmailStoreTableQuery.MaxItemsScanned"/>.</summary>
    public bool ScanLimitReached { get; }

    /// <summary>Actual read/sort/projection plan.</summary>
    public EmailStoreQueryPlan Plan { get; }
}
