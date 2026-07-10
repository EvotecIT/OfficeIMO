#nullable enable

using System.Globalization;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Infers a schema from the document header and sampled row values.
    /// </summary>
    /// <param name="sampleSize">Maximum number of rows to inspect. Defaults to 1000.</param>
    /// <returns>A schema containing one inferred column for each header field.</returns>
    public CsvSchema InferSchema(int sampleSize = 1000)
    {
        if (sampleSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(sampleSize), "Sample size must be greater than zero.");
        }

        using var rows = EnumerateRawRows().GetEnumerator();
        return InferSchema(rows, sampleSize, sampledRows: null);
    }

    private CsvSchema InferSchema(
        IEnumerator<object?[]> rows,
        int sampleSize,
        ICollection<object?[]>? sampledRows,
        bool cloneSampledRows = false)
    {
        var columns = new InferredColumn[_header.Count];
        for (var i = 0; i < _header.Count; i++)
        {
            columns[i] = new InferredColumn(_header[i]);
        }

        var sampledRowCount = 0;
        while (sampledRowCount < sampleSize && rows.MoveNext())
        {
            var row = rows.Current;
            if (sampledRows is not null)
            {
                sampledRows.Add(cloneSampledRows ? (object?[])row.Clone() : row);
            }

            for (var i = 0; i < columns.Length; i++)
            {
                var value = i < row.Length ? row[i] : null;
                columns[i].Observe(value, _culture, _dateTimeFormats);
            }

            sampledRowCount++;
        }

        var schemaColumns = new List<CsvSchemaColumn>(columns.Length);
        foreach (var column in columns)
        {
            schemaColumns.Add(column.ToSchemaColumn(sampledRowCount));
        }

        return new CsvSchema(schemaColumns);
    }

#if NET8_0_OR_GREATER
    private CsvSchema InferSchema(
        CsvParser.CsvTextDataReaderRowSource rows,
        int sampleSize,
        string? nullValue)
    {
        var columns = new InferredColumn[_header.Count];
        for (var i = 0; i < _header.Count; i++)
        {
            columns[i] = new InferredColumn(_header[i]);
        }

        var sampledRowCount = 0;
        while (sampledRowCount < sampleSize && rows.Read())
        {
            for (var i = 0; i < columns.Length; i++)
            {
                if (nullValue is not null && rows.IsNull(i, nullValue))
                {
                    columns[i].ObserveMissing();
                }
                else
                {
                    columns[i].Observe(rows.GetSpan(i), _culture, _dateTimeFormats);
                }
            }

            sampledRowCount++;
        }

        var schemaColumns = new List<CsvSchemaColumn>(columns.Length);
        foreach (var column in columns)
        {
            schemaColumns.Add(column.ToSchemaColumn(sampledRowCount));
        }

        return new CsvSchema(schemaColumns);
    }
#endif

    /// <summary>
    /// Infers and attaches a schema to the document so subsequent validation uses the sampled structure.
    /// </summary>
    /// <param name="sampleSize">Maximum number of rows to inspect. Defaults to 1000.</param>
    /// <returns>The current document.</returns>
    public CsvDocument EnsureInferredSchema(int sampleSize = 1000)
    {
        _schema = InferSchema(sampleSize);
        return this;
    }

    private sealed class InferredColumn
    {
        private bool _hasMissingValue;
        private bool _hasObservedValue;
        private bool _canInt32 = true;
        private bool _canInt64 = true;
        private bool _canDecimal = true;
        private bool _canDouble = true;
        private bool _canBoolean = true;
        private bool _canDateTime = true;

        public InferredColumn(string name)
        {
            Name = name;
        }

        public string Name { get; }

        public void Observe(object? value, CultureInfo culture, IReadOnlyList<string>? dateTimeFormats)
        {
            if (value is null || value is string { Length: 0 })
            {
                _hasMissingValue = true;
                return;
            }

            _hasObservedValue = true;
            if (value is string text)
            {
                ObserveString(text, culture, dateTimeFormats);
                return;
            }

            ObserveTyped(value);
        }

        public void ObserveMissing()
        {
            _hasMissingValue = true;
        }

#if NET8_0_OR_GREATER
        public void Observe(ReadOnlySpan<char> value, CultureInfo culture, IReadOnlyList<string>? dateTimeFormats)
        {
            if (value.Length == 0)
            {
                ObserveMissing();
                return;
            }

            _hasObservedValue = true;
            var matchedNumeric = false;
            if (_canInt32)
            {
                _canInt32 = int.TryParse(value, NumberStyles.Any, culture, out _);
                matchedNumeric = _canInt32;
            }

            if (_canInt64)
            {
                if (!matchedNumeric)
                {
                    _canInt64 = long.TryParse(value, NumberStyles.Any, culture, out _);
                    matchedNumeric = _canInt64;
                }
            }

            if (_canDecimal)
            {
                if (!matchedNumeric)
                {
                    _canDecimal = decimal.TryParse(value, NumberStyles.Any, culture, out _);
                    matchedNumeric = _canDecimal;
                }
            }

            if (_canDouble)
            {
                if (!matchedNumeric)
                {
                    _canDouble = double.TryParse(value, NumberStyles.Any, culture, out _);
                }
            }

            if (_canBoolean)
            {
                _canBoolean = bool.TryParse(value, out _);
            }

            if (_canDateTime)
            {
                _canDateTime = TryParseDateTime(value, culture, dateTimeFormats);
            }
        }
#endif

        public CsvSchemaColumn ToSchemaColumn(int sampledRows)
        {
            return new CsvSchemaColumn(Name)
            {
                DataType = ResolveDataType(),
                IsRequired = sampledRows > 0 && _hasObservedValue && !_hasMissingValue
            };
        }

        private void ObserveString(string text, CultureInfo culture, IReadOnlyList<string>? dateTimeFormats)
        {
            var matchedNumeric = false;
            if (_canInt32)
            {
                _canInt32 = int.TryParse(text, NumberStyles.Any, culture, out _);
                matchedNumeric = _canInt32;
            }

            if (_canInt64)
            {
                if (!matchedNumeric)
                {
                    _canInt64 = long.TryParse(text, NumberStyles.Any, culture, out _);
                    matchedNumeric = _canInt64;
                }
            }

            if (_canDecimal)
            {
                if (!matchedNumeric)
                {
                    _canDecimal = decimal.TryParse(text, NumberStyles.Any, culture, out _);
                    matchedNumeric = _canDecimal;
                }
            }

            if (_canDouble)
            {
                if (!matchedNumeric)
                {
                    _canDouble = double.TryParse(text, NumberStyles.Any, culture, out _);
                }
            }

            if (_canBoolean)
            {
                _canBoolean = IsBooleanText(text);
            }

            if (_canDateTime)
            {
                _canDateTime = TryParseDateTime(text, culture, dateTimeFormats);
            }
        }

        private void ObserveTyped(object value)
        {
            var type = Nullable.GetUnderlyingType(value.GetType()) ?? value.GetType();

            var canInt32 = type == typeof(byte) || type == typeof(short) || type == typeof(int);
            var canInt64 = canInt32 || type == typeof(long);
            var canDecimal = canInt64 || type == typeof(decimal);
            var canDouble = canDecimal || type == typeof(float) || type == typeof(double);
            var canBoolean = type == typeof(bool);
            var canDateTime = type == typeof(DateTime);

            _canInt32 &= canInt32;
            _canInt64 &= canInt64;
            _canDecimal &= canDecimal;
            _canDouble &= canDouble;
            _canBoolean &= canBoolean;
            _canDateTime &= canDateTime;
        }

        private Type ResolveDataType()
        {
            if (!_hasObservedValue)
            {
                return typeof(string);
            }

            if (_canInt32)
            {
                return typeof(int);
            }

            if (_canInt64)
            {
                return typeof(long);
            }

            if (_canDecimal)
            {
                return typeof(decimal);
            }

            if (_canDouble)
            {
                return typeof(double);
            }

            if (_canBoolean)
            {
                return typeof(bool);
            }

            return _canDateTime ? typeof(DateTime) : typeof(string);
        }

        private static bool IsBooleanText(string text) =>
            string.Equals(text, "true", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(text, "false", StringComparison.OrdinalIgnoreCase);

        private static bool TryParseDateTime(string text, CultureInfo culture, IReadOnlyList<string>? dateTimeFormats)
        {
            if (dateTimeFormats is { Count: > 0 } &&
                DateTime.TryParseExact(text, dateTimeFormats as string[] ?? dateTimeFormats.ToArray(), culture, DateTimeStyles.None, out _))
            {
                return true;
            }

            return DateTime.TryParse(text, culture, DateTimeStyles.None, out _);
        }

#if NET8_0_OR_GREATER
        private static bool TryParseDateTime(ReadOnlySpan<char> text, CultureInfo culture, IReadOnlyList<string>? dateTimeFormats)
        {
            if (dateTimeFormats is { Count: > 0 } &&
                DateTime.TryParseExact(text, dateTimeFormats as string[] ?? dateTimeFormats.ToArray(), culture, DateTimeStyles.None, out _))
            {
                return true;
            }

            return DateTime.TryParse(text, culture, DateTimeStyles.None, out _);
        }
#endif
    }
}
