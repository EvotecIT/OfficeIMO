#nullable enable
#if NET8_0_OR_GREATER
using System.Buffers;

namespace OfficeIMO.CSV;

/// <summary>Owns decoded field text and metadata while the source reader advances.</summary>
internal sealed class CsvRecordTextBatch : IDisposable
{
    private const int MaximumPooledTextLength = 64 * 1024;
    private readonly string?[] _values;
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
        FieldCount = fieldCount;
        Culture = culture;
        DateTimeFormats = dateFormats;
        int length = Math.Max(1, checked(RowCapacity * fieldCount));
        _values = ArrayPool<string?>.Shared.Rent(length);
        try { _fields = ArrayPool<CapturedField>.Shared.Rent(length); }
        catch { ArrayPool<string?>.Shared.Return(_values, clearArray: true); throw; }
    }

    internal int RowCapacity { get; }
    internal int FieldCount { get; }
    internal int Count { get; private set; }
    internal CultureInfo Culture { get; }
    internal IReadOnlyList<string>? DateTimeFormats { get; }

    internal void CaptureRow(CsvDataReader reader)
    {
        int offset = Count * FieldCount;
        for (int ordinal = 0; ordinal < FieldCount; ordinal++)
        {
            bool missing = reader.IsCurrentFieldMissing(ordinal);
            ReadOnlySpan<char> value = reader.GetCurrentSourceSpan(ordinal, out string? materialized);
            byte flags = (byte)((missing ? 1 : 0) | (!missing && reader.IsCurrentFieldNull(ordinal) ? 2 : 0));
            int start = _textLength;
            if (materialized is null)
            {
                if (value.IsEmpty) materialized = string.Empty;
                else if (EnsureTextCapacity(value.Length))
                {
                    value.CopyTo(_text.AsSpan(_textLength));
                    _textLength += value.Length;
                }
                else materialized = value.ToString();
            }
            _values[offset + ordinal] = materialized;
            _fields[offset + ordinal] = new CapturedField(start, value.Length, flags);
        }
        Count++;
    }

    // Large records keep an owned string instead of expanding pooled storage without a bound.
    // This is a storage choice, not a field-size limit.
    private bool EnsureTextCapacity(int fieldLength)
    {
        if (fieldLength > MaximumPooledTextLength - _textLength) return false;
        int required = _textLength + fieldLength;
        if (_text is not null && required <= _text.Length) return true;
        int capacity = Math.Min(MaximumPooledTextLength, Math.Max(required, (_text?.Length ?? 2048) * 2));
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
        return _row < Count;
    }

    internal string GetString(int ordinal)
    {
        int index = GetIndex(ordinal);
        if (_values[index] is { } value) return value;
        var field = _fields[index];
        return _values[index] = _text.AsSpan(field.Start, field.Length).ToString();
    }
    internal ReadOnlySpan<char> GetSpan(int ordinal)
    {
        int index = GetIndex(ordinal);
        if (_values[index] is { } value) return value.AsSpan();
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
        ArrayPool<string?>.Shared.Return(_values, clearArray: true);
        ArrayPool<CapturedField>.Shared.Return(_fields);
        if (_text is not null) ArrayPool<char>.Shared.Return(_text, clearArray: true);
    }
}
#endif
