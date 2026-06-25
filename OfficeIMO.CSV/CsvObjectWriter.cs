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
    private readonly bool _leaveOpen;
    private readonly StringBuilder _rowBuffer = new(1024);
    private IReadOnlyList<string>? _columns;
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
        _useDefaultWritePath = _options.FormulaInjectionPolicy == CsvFormulaInjectionPolicy.Preserve
            && _options.QuoteMode == CsvQuoteMode.AsNeeded
            && _quoteFields == null;
        _useAlwaysQuotedWritePath = _options.FormulaInjectionPolicy == CsvFormulaInjectionPolicy.Preserve
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

        var itemColumns = ObjectDataHelpers.GetColumnNames(item);
        EnsureColumns(itemColumns, requireOrder: !ObjectDataHelpers.IsDictionaryLike(item));
        var columns = _columns!;

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

        EnsureColumns(columns);
        if (values.Count != _columns!.Count)
        {
            throw new CsvException($"Row contains {values.Count} values but header defines {_columns.Count} columns.");
        }

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
            CsvWriter.WriteRecordBuffered(_writer, _rowBuffer, values, _options.Delimiter, _options.NewLine, _options.Culture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns);
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

        EnsureColumns(columns);
        if (values.Length != _columns!.Count)
        {
            throw new CsvException($"Row contains {values.Length} values but header defines {_columns.Count} columns.");
        }

        WriteBuffered(values);
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

        EnsureColumns(columns);
        if (valueCount != _columns!.Count)
        {
            throw new CsvException($"Row contains {valueCount} values but header defines {_columns.Count} columns.");
        }

        CsvWriter.WriteRecordBuffered(_writer, _rowBuffer, valueCount, state, valueAccessor, _options.Delimiter, _options.NewLine, _options.Culture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns);
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

        CsvWriter.WriteRecordBuffered(_writer, _rowBuffer, values, _options.Delimiter, _options.NewLine, _options.Culture, _options.FormulaInjectionPolicy, _options.QuoteMode, _quoteFields, _columns);
    }
}
