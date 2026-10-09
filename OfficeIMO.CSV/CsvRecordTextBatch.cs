#nullable enable
#if NET8_0_OR_GREATER
using System.Buffers;

namespace OfficeIMO.CSV;

/// <summary>Owns bounded decoded field references while the source reader advances.</summary>
internal sealed class CsvRecordTextBatch : IDisposable
{
    private readonly string?[] _values;
    private readonly byte[] _flags;
    private int _row = -1;
    private bool _disposed;

    internal CsvRecordTextBatch(int preferredRows, int fieldCount, CultureInfo culture, IReadOnlyList<string>? dateFormats)
    {
        const int maximumElements = 4 * 1024 * 1024;
        RowCapacity = fieldCount == 0 ? 1 : Math.Max(1, Math.Min(preferredRows, maximumElements / fieldCount));
        FieldCount = fieldCount;
        Culture = culture;
        DateTimeFormats = dateFormats;
        int length = Math.Max(1, checked(RowCapacity * fieldCount));
        _values = ArrayPool<string?>.Shared.Rent(length);
        try { _flags = ArrayPool<byte>.Shared.Rent(length); }
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
            _values[offset + ordinal] = reader.GetCurrentSourceString(ordinal);
            _flags[offset + ordinal] = (byte)((missing ? 1 : 0) | (!missing && reader.IsDBNull(ordinal) ? 2 : 0));
        }
        Count++;
    }

    internal bool Read()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        _row++;
        return _row < Count;
    }

    internal string GetString(int ordinal) => _values[GetIndex(ordinal)] ?? string.Empty;
    internal ReadOnlySpan<char> GetSpan(int ordinal) => GetString(ordinal).AsSpan();
    internal bool IsMissing(int ordinal) => (_flags[GetIndex(ordinal)] & 1) != 0;
    internal bool IsConfiguredNull(int ordinal) => (_flags[GetIndex(ordinal)] & 2) != 0;

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
        ArrayPool<byte>.Shared.Return(_flags, clearArray: true);
    }
}
#endif
