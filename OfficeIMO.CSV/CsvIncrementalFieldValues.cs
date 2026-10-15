#nullable enable
#if NET8_0_OR_GREATER
namespace OfficeIMO.CSV;

/// <summary>Field ranges borrowed from one incremental transport buffer until its next advance.</summary>
internal sealed class CsvIncrementalFieldValues : IReadOnlyList<string>
{
    private char[] _buffer = Array.Empty<char>();
    private (int Start, int Length)[] _ranges = new (int, int)[16];
    private string?[] _strings = new string?[16];
    public int Count { get; private set; }
    public string this[int index] => _strings[index] ??= GetSpan(index).ToString();

    internal void SetFields(char[] buffer, int start, int length, char delimiter, bool trim)
    {
        Clear();
        _buffer = buffer;
        var fields = new CsvUnquotedFieldEnumerator(buffer, start, length, delimiter, trim);
        while (fields.MoveNext())
        {
            AddField(fields.Start, fields.Length);
        }
    }

    private void AddField(int start, int length)
    {
        if (Count == _ranges.Length)
        {
            Array.Resize(ref _ranges, checked(Count * 2));
            Array.Resize(ref _strings, _ranges.Length);
        }
        _ranges[Count++] = (start, length);
    }

    internal ReadOnlySpan<char> GetSpan(int index)
    {
        if ((uint)index >= (uint)Count) throw new IndexOutOfRangeException();
        var range = _ranges[index];
        return _buffer.AsSpan(range.Start, range.Length);
    }

    internal string? GetMaterializedString(int index) => _strings[index];

    internal void Clear()
    {
        Array.Clear(_strings, 0, Count);
        Count = 0;
        _buffer = Array.Empty<char>();
    }

    public IEnumerator<string> GetEnumerator()
    {
        for (int index = 0; index < Count; index++) yield return this[index];
    }
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
#endif
