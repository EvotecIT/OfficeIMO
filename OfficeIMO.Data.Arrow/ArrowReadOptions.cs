namespace OfficeIMO.Data.Arrow;

/// <summary>Controls bounded conversion from a data reader to Apache Arrow record batches.</summary>
public sealed class ArrowReadOptions {
    private int _batchSize = 65_536;
    private int _decimalPrecision = 29;
    private int _decimalScale = 10;

    /// <summary>Gets or sets the maximum number of rows in one record batch.</summary>
    public int BatchSize {
        get => _batchSize;
        set => _batchSize = value > 0
            ? value
            : throw new ArgumentOutOfRangeException(nameof(value), "Batch size must be greater than zero.");
    }

    /// <summary>Gets or sets the precision used for CLR <see cref="decimal"/> columns.</summary>
    public int DecimalPrecision {
        get => _decimalPrecision;
        set {
            if (value is < 1 or > 38) {
                throw new ArgumentOutOfRangeException(nameof(value), "Decimal precision must be between 1 and 38.");
            }
            _decimalPrecision = value;
        }
    }

    /// <summary>Gets or sets the scale used for CLR <see cref="decimal"/> columns.</summary>
    public int DecimalScale {
        get => _decimalScale;
        set => _decimalScale = value is >= 0 and <= 38
            ? value
            : throw new ArgumentOutOfRangeException(nameof(value), "Decimal scale must be between zero and 38.");
    }

    /// <summary>
    /// Gets or sets whether CLR types without a native adapter are converted to invariant text.
    /// When false, encountering such a column throws <see cref="NotSupportedException"/>.
    /// </summary>
    public bool ConvertUnsupportedTypesToString { get; set; } = true;

    /// <summary>
    /// Gets or sets an optional CLR type for every source column, in ordinal order.
    /// When supplied, these types replace <see cref="System.Data.Common.DbDataReader.GetFieldType(int)"/>
    /// and allow callers to skip source-reader schema inference.
    /// </summary>
    public IReadOnlyList<Type>? ColumnTypes { get; set; }

    internal Type[]? ValidateAndSnapshotColumnTypes(int fieldCount) {
        if (_decimalScale > _decimalPrecision) {
            throw new ArgumentOutOfRangeException(
                nameof(DecimalScale),
                "Decimal scale cannot be greater than the configured precision.");
        }

        if (ColumnTypes == null) {
            return null;
        }
        if (ColumnTypes.Count != fieldCount) {
            throw new ArgumentException(
                $"Explicit Arrow column types contain {ColumnTypes.Count} entries for a {fieldCount}-column reader.",
                nameof(ColumnTypes));
        }

        var snapshot = new Type[fieldCount];
        for (int ordinal = 0; ordinal < snapshot.Length; ordinal++) {
            snapshot[ordinal] = ColumnTypes[ordinal]
                ?? throw new ArgumentException(
                    $"Explicit Arrow column type {ordinal} is null.",
                    nameof(ColumnTypes));
        }
        return snapshot;
    }
}
