#nullable enable

using System.Globalization;
using System.Text;
using OfficeIMO.Shared;

namespace OfficeIMO.CSV;

/// <summary>
/// Streams object rows to CSV without materializing a <see cref="CsvDocument"/>.
/// </summary>
public sealed class CsvObjectWriter : IDisposable
{
    private readonly TextWriter _writer;
    private readonly CsvSaveOptions _options;
    private readonly HashSet<string>? _quoteFields;
    private readonly bool _useDefaultWritePath;
    private readonly bool _useAlwaysQuotedWritePath;
    private readonly bool _useFormattedValueOptions;
    private readonly bool _leaveOpen;
    private readonly StringBuilder _rowBuffer = new(1024);
    private const int WideTextRowThreshold = 20;
    private IReadOnlyList<string>? _columns;
    private Func<object, object?[], bool>? _propertyProjector;
    private Func<object, string?[], CultureInfo, bool>? _propertyTextProjector;
    private object?[]? _propertyValues;
    private string?[]? _propertyTextValues;
    private bool _disposed;

    /// <summary>
    /// Initializes a streaming object writer.
    /// </summary>
    /// <param name="writer">Destination text writer.</param>
    /// <param name="options">Optional CSV save options.</param>
    /// <param name="leaveOpen">Whether to leave the destination writer open when this writer is disposed.</param>
    public CsvObjectWriter(TextWriter writer, CsvSaveOptions? options = null, bool leaveOpen = false)
    {
        _writer = writer ?? throw new ArgumentNullException(nameof(writer));
        _options = options ?? new CsvSaveOptions();
        _quoteFields = CsvWriter.CreateQuoteFieldSet(_options.QuoteFields);
        _useFormattedValueOptions = _options.NullValue is not null || _options.DateTimeFormat is not null || _options.UseUtc;
        _useDefaultWritePath = !_useFormattedValueOptions
            && _options.FormulaInjectionPolicy == CsvFormulaInjectionPolicy.Preserve
            && _options.QuoteMode == CsvQuoteMode.AsNeeded
            && _quoteFields == null;
        _useAlwaysQuotedWritePath = !_useFormattedValueOptions
            && _options.FormulaInjectionPolicy == CsvFormulaInjectionPolicy.Preserve
            && _options.QuoteMode == CsvQuoteMode.Always
            && _quoteFields == null;
        _leaveOpen = leaveOpen;
    }

    /// <summary>
    /// Gets whether at least one row was written.
    /// </summary>
    public bool HasRows => _columns != null;

    /// <summary>
    /// Writes one object row, inferring columns from the first row.
    /// </summary>
    /// <param name="item">Object, dictionary, or flattened row to write.</param>
    public void WriteObject(object? item)
    {
        ThrowIfDisposed();
        if (item == null)
        {
            if (_columns == null)
            {
                throw new ArgumentException("Data rows cannot be null.", nameof(item));
            }

            throw new InvalidOperationException("Data rows cannot contain null entries.");
        }

        if (_columns != null &&
            _propertyTextProjector != null &&
            _propertyTextValues != null &&
            _propertyTextProjector(item, _propertyTextValues, _options.Culture))
        {
            WriteTextBuffered(_propertyTextValues);
            return;
        }

        if (_columns != null &&
            _propertyProjector != null &&
            _propertyValues != null &&
            _propertyProjector(item, _propertyValues))
        {
            WriteBuffered(_propertyValues);
            return;
        }

        var itemColumns = ObjectDataHelpers.GetColumnNames(item);
        var dictionaryLike = ObjectDataHelpers.IsDictionaryLike(item);
        EnsureColumns(itemColumns, requireOrder: !dictionaryLike);
        var columns = _columns!;

        if (_useDefaultWritePath &&
            !dictionaryLike &&
            ObjectDataHelpers.TryCreatePropertyTextProjector(item, columns, out var textProjector))
        {
            _propertyTextProjector = textProjector;
            _propertyTextValues = new string?[columns.Count];
            textProjector!(item, _propertyTextValues, _options.Culture);
            WriteTextBuffered(_propertyTextValues);
            return;
        }

        if (!dictionaryLike &&
            ObjectDataHelpers.TryCreatePropertyProjector(item, columns, out var projector))
        {
            _propertyProjector = projector;
            _propertyValues = new object?[columns.Count];
            projector!(item, _propertyValues);
            WriteBuffered(_propertyValues);
            return;
        }

        var values = new object?[columns.Count];
        for (var i = 0; i < columns.Count; i++)
        {
            values[i] = ObjectDataHelpers.GetValue(item, columns[i]);
        }

        WriteBuffered(values);
    }

    /// <summary>
    /// Writes one already-projected row using the provided column order.
    /// </summary>
    /// <param name="columns">Column names for the first row and validation for later rows.</param>
    /// <param name="values">Values in the same order as <paramref name="columns"/>.</param>
    public void WriteRow(IReadOnlyList<string> columns, IReadOnlyList<object?> values)
    {
        ThrowIfDisposed();
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        if (values == null)
        {
            throw new ArgumentNullException(nameof(values));
        }

        ValidateProjectedValueCount(columns, values.Count);
        EnsureColumns(columns);

        if (_useDefaultWritePath && values is object?[] arrayValues)
        {
            CsvWriter.WriteRecordBufferedDefault(_writer, _rowBuffer, arrayValues, _options.Delimiter, _options.NewLine, _options.Culture);
        }
        else if (_useAlwaysQuotedWritePath && values is object?[] alwaysQuotedArrayValues)
        {
            CsvWriter.WriteRecordBufferedAlwaysQuoted(_writer, _rowBuffer, alwaysQuotedArrayValues, _options.Delimiter, _options.NewLine, _options.Culture);
        }
        else
        {
            CsvWriter.WriteRecordBuffered(_writer, _rowBuffer, values, _options.Delimiter, _options.NewLine, _options.Culture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns, _options.DateTimeFormat, _options.UseUtc, _options.NullValue);
        }
    }

    /// <summary>
    /// Writes one already-projected row using the provided column order.
    /// </summary>
    /// <param name="columns">Column names for the first row and validation for later rows.</param>
    /// <param name="values">Values in the same order as <paramref name="columns"/>.</param>
    public void WriteRow(IReadOnlyList<string> columns, object?[] values)
    {
        ThrowIfDisposed();
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        if (values == null)
        {
            throw new ArgumentNullException(nameof(values));
        }

        ValidateProjectedValueCount(columns, values.Length);
        EnsureColumns(columns);

        WriteBuffered(values);
    }

    /// <summary>
    /// Writes one already-formatted text row using the provided column order.
    /// </summary>
    /// <param name="columns">Column names for the first row and validation for later rows.</param>
    /// <param name="values">Text values in the same order as <paramref name="columns"/>.</param>
    /// <remarks>
    /// Use this when the caller already owns culture-aware value formatting.
    /// The method still applies CSV escaping and validates that the row width matches the header.
    /// </remarks>
    public void WriteTextRow(IReadOnlyList<string> columns, string?[] values)
    {
        ThrowIfDisposed();
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        if (values == null)
        {
            throw new ArgumentNullException(nameof(values));
        }

        ValidateProjectedValueCount(columns, values.Length);
        EnsureColumns(columns);

        WriteTextBuffered(values);
    }

    /// <summary>
    /// Writes one already-projected row using the column order that was established by a previous row.
    /// </summary>
    /// <param name="values">Values in the same order as the established CSV columns.</param>
    /// <remarks>
    /// Use this only when the caller already owns schema validation and can guarantee stable column order.
    /// The method still validates that the row width matches the established header.
    /// </remarks>
    public void WriteTrustedRow(object?[] values)
    {
        ThrowIfDisposed();
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(values);
#else
        if (values == null)
        {
            throw new ArgumentNullException(nameof(values));
        }
#endif

        if (_columns == null)
        {
            throw new InvalidOperationException("Columns must be established before writing trusted rows.");
        }

        if (values.Length != _columns.Count)
        {
            throw new CsvException($"Row contains {values.Length} values but header defines {_columns.Count} columns.");
        }

        WriteBuffered(values);
    }

    /// <summary>
    /// Writes one already-projected row using a caller-provided value accessor and the established column order.
    /// </summary>
    /// <typeparam name="TState">State type used by the value accessor.</typeparam>
    /// <param name="valueCount">Number of values exposed by <paramref name="valueAccessor"/>.</param>
    /// <param name="state">Caller state passed to <paramref name="valueAccessor"/> for each column index.</param>
    /// <param name="valueAccessor">Function that returns the value for a column index.</param>
    public void WriteTrustedRow<TState>(
        int valueCount,
        TState state,
        Func<TState, int, object?> valueAccessor)
    {
        ThrowIfDisposed();
        if (valueAccessor == null)
        {
            throw new ArgumentNullException(nameof(valueAccessor));
        }

        if (_columns == null)
        {
            throw new InvalidOperationException("Columns must be established before writing trusted rows.");
        }

        if (valueCount != _columns.Count)
        {
            throw new CsvException($"Row contains {valueCount} values but header defines {_columns.Count} columns.");
        }

        WriteBuffered(valueCount, state, valueAccessor);
    }

    /// <summary>
    /// Writes one already-formatted text row using the column order that was established by a previous row.
    /// </summary>
    /// <param name="values">Text values in the same order as the established CSV columns.</param>
    /// <remarks>
    /// Use this only when the caller already owns schema validation and culture-aware value formatting.
    /// The method still applies CSV escaping and validates that the row width matches the established header.
    /// </remarks>
    public void WriteTrustedTextRow(string?[] values)
    {
        ThrowIfDisposed();
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(values);
#else
        if (values == null)
        {
            throw new ArgumentNullException(nameof(values));
        }
#endif

        if (_columns == null)
        {
            throw new InvalidOperationException("Columns must be established before writing trusted rows.");
        }

        if (values.Length != _columns.Count)
        {
            throw new CsvException($"Row contains {values.Length} values but header defines {_columns.Count} columns.");
        }

        if (_useDefaultWritePath)
        {
            WriteDefaultTextRecord(values);
            return;
        }

        if (_useAlwaysQuotedWritePath)
        {
            CsvWriter.WriteRecordBufferedAlwaysQuoted(_writer, _rowBuffer, values, _options.Delimiter, _options.NewLine);
            return;
        }

        CsvWriter.WriteRecordBuffered<string?>(_writer, _rowBuffer, values, _options.Delimiter, _options.NewLine, _options.Culture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns, _options.DateTimeFormat, _options.UseUtc, _options.NullValue);
    }

    /// <summary>
    /// Writes one already-formatted text row using a caller-provided value accessor.
    /// </summary>
    /// <typeparam name="TState">State type used by the value accessor.</typeparam>
    /// <param name="columns">Column names for the first row and validation for later rows.</param>
    /// <param name="valueCount">Number of values exposed by <paramref name="valueAccessor"/>.</param>
    /// <param name="state">Caller state passed to <paramref name="valueAccessor"/> for each column index.</param>
    /// <param name="valueAccessor">Function that returns the already-formatted text for a column index.</param>
    public void WriteTextRow<TState>(
        IReadOnlyList<string> columns,
        int valueCount,
        TState state,
        Func<TState, int, string?> valueAccessor)
    {
        ThrowIfDisposed();
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        if (valueAccessor == null)
        {
            throw new ArgumentNullException(nameof(valueAccessor));
        }

        ValidateProjectedValueCount(columns, valueCount);
        EnsureColumns(columns);

        WriteTextBuffered(valueCount, state, valueAccessor);
    }

    /// <summary>
    /// Writes one already-formatted text row using the column order that was established by a previous row.
    /// </summary>
    /// <typeparam name="TState">State type used by the value accessor.</typeparam>
    /// <param name="valueCount">Number of values exposed by <paramref name="valueAccessor"/>.</param>
    /// <param name="state">Caller state passed to <paramref name="valueAccessor"/> for each column index.</param>
    /// <param name="valueAccessor">Function that returns the already-formatted text for a column index.</param>
    public void WriteTrustedTextRow<TState>(
        int valueCount,
        TState state,
        Func<TState, int, string?> valueAccessor)
    {
        ThrowIfDisposed();
        if (valueAccessor == null)
        {
            throw new ArgumentNullException(nameof(valueAccessor));
        }

        if (_columns == null)
        {
            throw new InvalidOperationException("Columns must be established before writing trusted rows.");
        }

        if (valueCount != _columns.Count)
        {
            throw new CsvException($"Row contains {valueCount} values but header defines {_columns.Count} columns.");
        }

        WriteTextBuffered(valueCount, state, valueAccessor);
    }

    /// <summary>
    /// Writes one projected row using a caller-provided value accessor.
    /// </summary>
    /// <typeparam name="TState">State type used by the value accessor.</typeparam>
    /// <param name="columns">Column names for the first row and validation for later rows.</param>
    /// <param name="valueCount">Number of values exposed by <paramref name="valueAccessor"/>.</param>
    /// <param name="state">Caller state passed to <paramref name="valueAccessor"/> for each column index.</param>
    /// <param name="valueAccessor">Function that returns the value for a column index.</param>
    public void WriteRow<TState>(
        IReadOnlyList<string> columns,
        int valueCount,
        TState state,
        Func<TState, int, object?> valueAccessor)
    {
        ThrowIfDisposed();
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        if (valueAccessor == null)
        {
            throw new ArgumentNullException(nameof(valueAccessor));
        }

        ValidateProjectedValueCount(columns, valueCount);
        EnsureColumns(columns);

        WriteBuffered(valueCount, state, valueAccessor);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        if (_leaveOpen)
        {
            _writer.Flush();
        }
        else
        {
            _writer.Dispose();
        }
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(CsvObjectWriter));
        }
    }

    private void EnsureColumns(IReadOnlyList<string> columns, bool requireOrder = true)
    {
        if (_columns != null)
        {
            ValidateColumns(columns, requireOrder);
            return;
        }

        if (columns.Count == 0)
        {
            throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries.");
        }

        _columns = columns.ToArray();
        if (_options.IncludeHeader)
        {
            WriteHeader();
        }
    }

    private void ValidateProjectedValueCount(IReadOnlyList<string> columns, int valueCount)
    {
        var expectedCount = _columns?.Count ?? columns.Count;
        if (valueCount != expectedCount)
        {
            throw new CsvException($"Row contains {valueCount} values but header defines {expectedCount} columns.");
        }
    }

    private void ValidateColumns(IReadOnlyList<string> columns, bool requireOrder)
    {
        if (columns.Count != _columns!.Count)
        {
            throw new CsvException($"Row defines {columns.Count} columns but header defines {_columns.Count} columns.");
        }

        if (!requireOrder)
        {
            ValidateColumnSet(columns);
            return;
        }

        for (var i = 0; i < columns.Count; i++)
        {
            if (!string.Equals(columns[i], _columns[i], StringComparison.Ordinal))
            {
                throw new CsvException($"Row column '{columns[i]}' at index {i} does not match header column '{_columns[i]}'.");
            }
        }
    }

    private void ValidateColumnSet(IReadOnlyList<string> columns)
    {
        var expected = new HashSet<string>(_columns!, StringComparer.Ordinal);
        for (var i = 0; i < columns.Count; i++)
        {
            if (!expected.Remove(columns[i]))
            {
                throw new CsvException($"Row column '{columns[i]}' does not match the header columns.");
            }
        }
    }

    private void WriteHeader()
    {
        if (_useDefaultWritePath)
        {
            CsvWriter.WriteRecord(_writer, _columns!, _options.Delimiter, _options.NewLine, CultureInfo.InvariantCulture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns);
            return;
        }

        if (_useAlwaysQuotedWritePath)
        {
            CsvWriter.WriteRecordBufferedAlwaysQuoted(_writer, _rowBuffer, _columns!, _options.Delimiter, _options.NewLine, CultureInfo.InvariantCulture);
            return;
        }

        CsvWriter.WriteRecord(_writer, _columns!, _options.Delimiter, _options.NewLine, CultureInfo.InvariantCulture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns);
    }

    private void WriteBuffered(object?[] values)
    {
        if (_useDefaultWritePath)
        {
            CsvWriter.WriteRecordBufferedDefault(_writer, _rowBuffer, values, _options.Delimiter, _options.NewLine, _options.Culture);
            return;
        }

        if (_useAlwaysQuotedWritePath)
        {
            CsvWriter.WriteRecordBufferedAlwaysQuoted(_writer, _rowBuffer, values, _options.Delimiter, _options.NewLine, _options.Culture);
            return;
        }

        CsvWriter.WriteRecordBuffered(_writer, _rowBuffer, values, _options.Delimiter, _options.NewLine, _options.Culture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns, _options.DateTimeFormat, _options.UseUtc, _options.NullValue);
    }

    private void WriteBuffered<TState>(int valueCount, TState state, Func<TState, int, object?> valueAccessor)
    {
        if (_useDefaultWritePath)
        {
            CsvWriter.WriteRecordBufferedDefault(_writer, _rowBuffer, valueCount, state, valueAccessor, _options.Delimiter, _options.NewLine, _options.Culture);
            return;
        }

        if (_useAlwaysQuotedWritePath)
        {
            CsvWriter.WriteRecordBuffered(_writer, _rowBuffer, valueCount, state, valueAccessor, _options.Delimiter, _options.NewLine, _options.Culture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns, _options.DateTimeFormat, _options.UseUtc, _options.NullValue);
            return;
        }

        CsvWriter.WriteRecordBuffered(_writer, _rowBuffer, valueCount, state, valueAccessor, _options.Delimiter, _options.NewLine, _options.Culture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns, _options.DateTimeFormat, _options.UseUtc, _options.NullValue);
    }

    private void WriteTextBuffered(string?[] values)
    {
        if (_useDefaultWritePath)
        {
            WriteDefaultTextRecord(values);
            return;
        }

        if (_useAlwaysQuotedWritePath)
        {
            CsvWriter.WriteRecordBufferedAlwaysQuoted(_writer, _rowBuffer, values, _options.Delimiter, _options.NewLine);
            return;
        }

        CsvWriter.WriteRecordBuffered<string?>(_writer, _rowBuffer, values, _options.Delimiter, _options.NewLine, _options.Culture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns, _options.DateTimeFormat, _options.UseUtc, _options.NullValue);
    }

    private void WriteTextBuffered<TState>(int valueCount, TState state, Func<TState, int, string?> valueAccessor)
    {
        if (_useDefaultWritePath)
        {
            CsvWriter.WriteRecordBufferedDefault(_writer, _rowBuffer, valueCount, state, valueAccessor, _options.Delimiter, _options.NewLine);
            return;
        }

        if (_useAlwaysQuotedWritePath)
        {
            CsvWriter.WriteRecordBufferedAlwaysQuoted(_writer, _rowBuffer, valueCount, state, valueAccessor, _options.Delimiter, _options.NewLine);
            return;
        }

        CsvWriter.WriteTextRecordBuffered(_writer, _rowBuffer, valueCount, state, valueAccessor, _options.Delimiter, _options.NewLine, _options.Culture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns, _options.DateTimeFormat, _options.UseUtc, _options.NullValue);
    }

    private void WriteDefaultTextRecord(string?[] values)
    {
        if (values.Length >= WideTextRowThreshold)
        {
            CsvWriter.WriteRecordDefault(_writer, values, _options.Delimiter, _options.NewLine);
            return;
        }

        CsvWriter.WriteRecordBufferedDefault(_writer, _rowBuffer, values, _options.Delimiter, _options.NewLine);
    }
}
