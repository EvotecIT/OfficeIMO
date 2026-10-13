#nullable enable
#if NET8_0_OR_GREATER
using System.Buffers;

namespace OfficeIMO.CSV;

/// <summary>Owns decoded field text and metadata while the source reader advances.</summary>
internal sealed partial class CsvRecordTextBatch : IDisposable
{
    private readonly int _maximumPooledTextLength;
    private string?[]? _values;
    private readonly CapturedField[] _fields;
    private char[]? _text;
    private int _textLength;
    private int _row = -1;
    private bool _disposed;

    private readonly struct CapturedField
    {
        internal CapturedField(int start, int length, byte flags)
        {
            Start = start;
            Length = length;
            Flags = flags;
        }
        internal int Start { get; }
        internal int Length { get; }
        internal byte Flags { get; }
    }

    internal CsvRecordTextBatch(int preferredRows, int fieldCount, CultureInfo culture, IReadOnlyList<string>? dateFormats)
    {
        const int maximumElements = 4 * 1024 * 1024;
        RowCapacity = fieldCount == 0 ? 1 : Math.Max(1, Math.Min(preferredRows, maximumElements / fieldCount));
        // Scale owned text with the effective batch size while retaining a fixed upper bound.
        _maximumPooledTextLength = Math.Max(256, Math.Min(RowCapacity, 1024)) * 256;
        FieldCount = fieldCount;
        Culture = culture;
        DateTimeFormats = dateFormats;
        int length = Math.Max(1, checked(RowCapacity * fieldCount));
        _fields = ArrayPool<CapturedField>.Shared.Rent(length);
    }

    internal int RowCapacity { get; }
    internal int FieldCount { get; }
    internal int Count { get; private set; }
    internal CultureInfo Culture { get; }
    internal IReadOnlyList<string>? DateTimeFormats { get; }

    /// <summary>Reuses owned storage after the prior worker and its ordered merge have completed.</summary>
    internal void Reset()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_values is not null) Array.Clear(_values, 0, Count * FieldCount);
        // Capture initializes metadata for every visible row; only cached strings retain objects.
        Count = 0;
        _row = -1;
        _textLength = 0;
        _rawDelimiter = default;
        _rawTrim = false;
        _rawNullValue = null;
    }

    internal void CaptureRow(CsvDataReader reader)
    {
        int offset = Count * FieldCount;
        for (int ordinal = 0; ordinal < FieldCount; ordinal++)
        {
            bool missing = reader.IsCurrentFieldMissing(ordinal);
            ReadOnlySpan<char> value = reader.GetCurrentSourceSpan(ordinal, out string? materialized);
            byte flags = (byte)((missing ? 1 : 0) | (!missing && reader.IsCurrentFieldNull(ordinal) ? 2 : 0));
            CaptureField(offset + ordinal, value, materialized, flags);
        }
        Count++;
    }

    /// <summary>Detaches already-aligned raw fields without repeating reader getter validation.</summary>
    internal void CaptureRow(CsvIncrementalRowSource source, string? nullValue)
    {
        int offset = Count * FieldCount;
        for (int ordinal = 0; ordinal < FieldCount; ordinal++)
        {
            bool missing = source.IsMissing(ordinal);
            ReadOnlySpan<char> value = source.GetSpan(ordinal, out string? materialized);
            bool isNull = !missing && nullValue is not null && value.SequenceEqual(nullValue.AsSpan());
            CaptureField(offset + ordinal, value, materialized, (byte)((missing ? 1 : 0) | (isNull ? 2 : 0)));
        }
        Count++;
    }

    private void CaptureField(int index, ReadOnlySpan<char> value, string? materialized, byte flags)
    {
        int start = _textLength;
        if (materialized is null && !value.IsEmpty)
        {
            if (EnsureTextCapacity(value.Length))
            {
                value.CopyTo(_text.AsSpan(_textLength));
                _textLength += value.Length;
            }
            else materialized = value.ToString();
        }
        if (materialized is { Length: > 0 }) EnsureValues()[index] = materialized;
        else if (_values is not null) _values[index] = materialized;
        _fields[index] = new CapturedField(start, value.Length, flags);
    }

    // Span-only batches need no string cache. Clear a late rental for all previously captured slots.
    private string?[] EnsureValues()
    {
        if (_values is not null) return _values;
        string?[] values = ArrayPool<string?>.Shared.Rent(Math.Max(1, checked(RowCapacity * FieldCount)));
        Array.Clear(values, 0, values.Length);
        return _values = values;
    }

    // Large records keep an owned string instead of expanding pooled storage without a bound.
    // This is a storage choice, not a field-size limit.
    private bool EnsureTextCapacity(int fieldLength)
    {
        if (fieldLength > _maximumPooledTextLength - _textLength) return false;
        int required = _textLength + fieldLength;
        if (_text is not null && required <= _text.Length) return true;
        int capacity = Math.Min(_maximumPooledTextLength, Math.Max(required, (_text?.Length ?? 2048) * 2));
        char[] next = ArrayPool<char>.Shared.Rent(capacity);
        if (_text is not null)
        {
            _text.AsSpan(0, _textLength).CopyTo(next);
            ArrayPool<char>.Shared.Return(_text, clearArray: true);
        }
        _text = next;
        return true;
    }

    internal bool Read()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        _row++;
        if (_row >= Count) return false;
        IndexRawRecord();
        return true;
    }

    internal string GetString(int ordinal)
    {
        int index = GetIndex(ordinal);
        if (_values is not null && _values[index] is { } value) return value;
        var field = _fields[index];
        if (field.Length == 0) return string.Empty;
        return EnsureValues()[index] = _text.AsSpan(field.Start, field.Length).ToString();
    }
    internal ReadOnlySpan<char> GetSpan(int ordinal)
    {
        int index = GetIndex(ordinal);
        if (_values is not null && _values[index] is { } value) return value.AsSpan();
        var field = _fields[index];
        return _text.AsSpan(field.Start, field.Length);
    }
    internal bool IsMissing(int ordinal) => (_fields[GetIndex(ordinal)].Flags & 1) != 0;
    internal bool IsConfiguredNull(int ordinal) => (_fields[GetIndex(ordinal)].Flags & 2) != 0;

    private int GetIndex(int ordinal)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if ((uint)ordinal >= (uint)FieldCount) throw new IndexOutOfRangeException();
        if ((uint)_row >= (uint)Count) throw new InvalidOperationException("The CSV batch is not positioned on a row.");
        return _row * FieldCount + ordinal;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        if (_values is not null) ArrayPool<string?>.Shared.Return(_values, clearArray: true);
        ArrayPool<CapturedField>.Shared.Return(_fields);
        if (_text is not null) ArrayPool<char>.Shared.Return(_text, clearArray: true);
    }
}
#endif
