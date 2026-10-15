#nullable enable
#if NET8_0_OR_GREATER
namespace OfficeIMO.CSV;

/// <summary>Enumerates unchanged field ranges in a complete, already-validated unquoted record.</summary>
internal ref struct CsvUnquotedFieldEnumerator
{
    private readonly ReadOnlySpan<char> _record;
    private readonly int _start;
    private readonly char _delimiter;
    private readonly bool _trim;
    private int _position;

    internal CsvUnquotedFieldEnumerator(char[] buffer, int start, int length, char delimiter, bool trim)
    {
        _record = buffer.AsSpan(start, length);
        _start = start;
        _delimiter = delimiter;
        _trim = trim;
        _position = 0;
        Start = Length = 0;
    }

    internal int Start { get; private set; }
    internal int Length { get; private set; }

    internal bool MoveNext()
    {
        if (_position < 0) return false;
        int start = _position;
        int separator = _record.Slice(start).IndexOf(_delimiter);
        int end = separator < 0 ? _record.Length : start + separator;
        _position = separator < 0 ? -1 : end + 1;
        if (_trim)
        {
            while (start < end && char.IsWhiteSpace(_record[start])) start++;
            while (end > start && char.IsWhiteSpace(_record[end - 1])) end--;
        }
        Start = _start + start;
        Length = end - start;
        return true;
    }
}
#endif
