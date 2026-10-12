#if NET8_0_OR_GREATER
#nullable enable

namespace OfficeIMO.CSV;

public sealed partial class CsvRowWriter
{
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);
    private byte[]? _utf8Delimiter;
    private ReadOnlyMemory<byte>?[]? _utf8Values;

    /// <summary>Writes one already-formatted UTF-8 row in the supplied column order.</summary>
    /// <param name="columns">Column names used to establish or validate the schema.</param>
    /// <param name="values">UTF-8 field text. Null uses NullValue; empty memory is an empty text field.</param>
    /// <remarks>
    /// All fields are validated before writing the row or its initial header. Invalid UTF-8 throws
    /// DecoderFallbackException. Borrowed memory is consumed synchronously and is not retained.
    /// CSV quoting and formula escaping still apply. UTF-8 stream and file destinations without a
    /// BOM copy field bytes directly; other encodings and supplied TextWriters decode the fields.
    /// </remarks>
    public void WriteUtf8Row(IReadOnlyList<string> columns, ReadOnlyMemory<byte>?[] values)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(columns);
        ArgumentNullException.ThrowIfNull(values);
        ValidateProjectedValueCount(columns, values.Length);
        for (int i = 0; i < values.Length; i++)
        {
            if (values[i] is { } value) StrictUtf8.GetCharCount(value.Span);
        }

        if (_writer is PooledUtf8TextWriter utf8Writer)
        {
            _utf8Delimiter ??= utf8Writer.Encoding.GetBytes(_delimiterText);
            EnsureColumns(columns);
            CsvWriter.WriteUtf8Record(utf8Writer, values, _utf8Delimiter, _options, _quoteFields, _columns!);
        }
        else
        {
            var text = new string?[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                text[i] = values[i] is { } value ? StrictUtf8.GetString(value.Span) : null;
            }
            EnsureColumns(columns);
            WriteTextBuffered(text);
        }
    }

    /// <summary>Writes UTF-8 text using the schema established by a previous row.</summary>
    /// <param name="values">Validated synchronously; null and empty memory retain their distinct formatting semantics.</param>
    public void WriteUtf8Row(ReadOnlyMemory<byte>?[] values)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(values);
        WriteUtf8Row(_columns ?? throw new InvalidOperationException("Columns must be established before writing rows without an explicit schema."), values);
    }

    /// <summary>Writes a UTF-8 row using a caller-provided accessor.</summary>
    /// <typeparam name="TState">Caller state type.</typeparam>
    /// <param name="columns">Column names used to establish or validate the schema.</param>
    /// <param name="valueCount">Number of fields returned by the accessor.</param>
    /// <param name="state">State passed to the accessor.</param>
    /// <param name="valueAccessor">Called once per field. Borrowed values are consumed before this method returns.</param>
    /// <remarks>Returned field memory must remain valid for the entire method call so all fields can be validated before output begins.</remarks>
    public void WriteUtf8Row<TState>(IReadOnlyList<string> columns, int valueCount, TState state, Func<TState, int, ReadOnlyMemory<byte>?> valueAccessor)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(columns);
        ArgumentNullException.ThrowIfNull(valueAccessor);
        ValidateProjectedValueCount(columns, valueCount);
        if (_utf8Values?.Length != valueCount) _utf8Values = new ReadOnlyMemory<byte>?[valueCount];
        try
        {
            for (int i = 0; i < valueCount; i++) _utf8Values[i] = valueAccessor(state, i);
            WriteUtf8Row(columns, _utf8Values);
        }
        finally
        {
            Array.Clear(_utf8Values);
        }
    }

    /// <summary>Writes UTF-8 text through an accessor using the established schema.</summary>
    /// <typeparam name="TState">Caller state type.</typeparam>
    /// <param name="valueCount">Number of fields returned by the accessor.</param>
    /// <param name="state">State passed to the accessor.</param>
    /// <param name="valueAccessor">Called once per field; its memory is consumed synchronously.</param>
    public void WriteUtf8Row<TState>(int valueCount, TState state, Func<TState, int, ReadOnlyMemory<byte>?> valueAccessor)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(valueAccessor);
        WriteUtf8Row(_columns ?? throw new InvalidOperationException("Columns must be established before writing rows without an explicit schema."), valueCount, state, valueAccessor);
    }
}
#endif
