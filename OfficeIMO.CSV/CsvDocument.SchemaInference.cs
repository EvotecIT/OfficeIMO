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
            if (_canInt32)
            {
                _canInt32 = int.TryParse(text, NumberStyles.Any, culture, out _);
            }

            if (_canInt64)
            {
                _canInt64 = long.TryParse(text, NumberStyles.Any, culture, out _);
            }

            if (_canDecimal)
            {
                _canDecimal = decimal.TryParse(text, NumberStyles.Any, culture, out _);
            }

            if (_canDouble)
            {
                _canDouble = double.TryParse(text, NumberStyles.Any, culture, out _);
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
    }
}
