#nullable enable
#if NET8_0_OR_GREATER
namespace OfficeIMO.CSV;

internal sealed partial class CsvRecordTextBatch
{
    private const byte RawRecord = 4;
    private char _rawDelimiter;
    private bool _rawTrim;
    private string? _rawNullValue;

    internal int RemainingRawTextLength => MaximumPooledTextLength - _textLength;
    internal int RawTextOffset => _textLength;

    internal void ConfigureRawRecords(char delimiter, bool trim, string? nullValue)
    {
        _rawDelimiter = delimiter;
        _rawTrim = trim;
        _rawNullValue = nullValue;
    }

    // The first field slot holds the record range until its worker splits this row.
    // Raw capture requires a nonempty header; canonical rows retain their usual field metadata.
    internal void CaptureRawRecord(int start, int length)
    {
        _fields[Count * FieldCount] = new CapturedField(start, length, RawRecord);
        Count++;
    }

    internal void AppendRawText(char[] buffer, int start, int length)
    {
        if (!EnsureTextCapacity(length))
            throw new InvalidOperationException("The CSV raw record block exceeds its reserved text capacity.");
        buffer.AsSpan(start, length).CopyTo(_text.AsSpan(_textLength));
        _textLength += length;
    }

    private void IndexRawRecord()
    {
        if (FieldCount == 0) return;
        int offset = _row * FieldCount;
        CapturedField record = _fields[offset];
        if ((record.Flags & RawRecord) == 0) return;

        var fields = new CsvUnquotedFieldEnumerator(_text!, record.Start, record.Length, _rawDelimiter, _rawTrim);
        int ordinal = 0;
        while (ordinal < FieldCount && fields.MoveNext())
        {
            ReadOnlySpan<char> value = _text.AsSpan(fields.Start, fields.Length);
            bool isNull = _rawNullValue is not null && value.SequenceEqual(_rawNullValue.AsSpan());
            int index = offset + ordinal++;
            // Rented arrays can contain arbitrary prior values. Initialize every visible slot.
            _values[index] = value.IsEmpty ? string.Empty : null;
            _fields[index] = new CapturedField(fields.Start, fields.Length, (byte)(isNull ? 2 : 0));
        }
        while (ordinal < FieldCount)
        {
            int index = offset + ordinal++;
            _values[index] = string.Empty;
            _fields[index] = new CapturedField(0, 0, 1);
        }
    }
}
#endif
