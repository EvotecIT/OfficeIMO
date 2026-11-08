#if NET8_0_OR_GREATER
using System.Buffers;
using System.Text;

namespace OfficeIMO.CSV;

internal sealed partial class CsvRecordTextBatch
{
    private const byte Utf8Field = 8;
    private readonly Encoding? _utf8Encoding;
    private byte[]? _utf8Bytes;
    private int _utf8Length;

    internal bool CanCaptureUtf8 => _utf8Encoding is not null;

    /// <summary>Detaches canonical UTF-8 fields after the reader has advanced and aligned the row.</summary>
    internal void CaptureRow(CsvParser.CsvUtf8StreamDataReaderRowSource source, string? nullValue)
    {
        int offset = Count * FieldCount;
        for (int ordinal = 0; ordinal < FieldCount; ordinal++)
        {
            bool missing = source.IsMissing(ordinal);
            bool isNull = !missing && source.IsNull(ordinal, nullValue);
            byte flags = (byte)((missing ? 1 : 0) | (isNull ? 2 : 0));
            if (source.TryGetUtf8Text(ordinal, out var bytes) &&
                TryCaptureUtf8Field(offset + ordinal, bytes, flags)) continue;
            string materialized = source.GetString(ordinal);
            CaptureField(offset + ordinal, materialized.AsSpan(), materialized, flags);
        }
        Count++;
    }

    // UTF-8 fields remain owned bytes until a worker actually needs decoded text.
    private bool TryCaptureUtf8Field(int index, ReadOnlySpan<byte> value, byte flags)
    {
        if (value.Length > _maximumPooledTextLength - _utf8Length) return false;
        int required = _utf8Length + value.Length;
        if (required > (_utf8Bytes?.Length ?? 0))
        {
            int capacity = Math.Min(_maximumPooledTextLength, Math.Max(required, (_utf8Bytes?.Length ?? 2048) * 2));
            byte[] next = ArrayPool<byte>.Shared.Rent(capacity);
            if (_utf8Bytes is not null)
            {
                _utf8Bytes.AsSpan(0, _utf8Length).CopyTo(next);
                ArrayPool<byte>.Shared.Return(_utf8Bytes, clearArray: true);
            }
            _utf8Bytes = next;
        }
        value.CopyTo(_utf8Bytes.AsSpan(_utf8Length));
        if (_values is not null) _values[index] = null;
        _fields[index] = new CapturedField(_utf8Length, value.Length, (byte)(flags | Utf8Field));
        _utf8Length = required;
        return true;
    }

    /// <summary>Decodes owned fields on their worker before the ordinary text reader consumes them.</summary>
    internal void PrepareDecodedText()
    {
        if (_utf8Length == 0) return;
        ReadOnlySpan<byte> bytes = _utf8Bytes.AsSpan(0, _utf8Length);
        int fields = Count * FieldCount;
        if (EnsureTextCapacity(_utf8Length))
        {
            if (Ascii.IsValid(bytes))
            {
                int first = _textLength;
                _textLength += _utf8Encoding!.GetChars(bytes, _text.AsSpan(_textLength));
                for (int i = 0; i < fields; i++)
                {
                    CapturedField field = _fields[i];
                    if ((field.Flags & Utf8Field) != 0)
                        _fields[i] = new CapturedField(first + field.Start, field.Length, (byte)(field.Flags & ~Utf8Field));
                }
            }
            else
            {
                for (int i = 0; i < fields; i++)
                {
                    CapturedField field = _fields[i];
                    if ((field.Flags & Utf8Field) == 0) continue;
                    int first = _textLength;
                    int length = _utf8Encoding!.GetChars(bytes.Slice(field.Start, field.Length), _text.AsSpan(first));
                    _textLength += length;
                    _fields[i] = new CapturedField(first, length, (byte)(field.Flags & ~Utf8Field));
                }
            }
        }
        else
        {
            // This is a bounded storage fallback, not a new input-size restriction.
            for (int i = 0; i < fields; i++)
            {
                CapturedField field = _fields[i];
                if ((field.Flags & Utf8Field) == 0) continue;
                string value = _utf8Encoding!.GetString(bytes.Slice(field.Start, field.Length));
                EnsureValues()[i] = value;
                _fields[i] = new CapturedField(0, value.Length, (byte)(field.Flags & ~Utf8Field));
            }
        }
        ArrayPool<byte>.Shared.Return(_utf8Bytes!, clearArray: true);
        _utf8Bytes = null;
        _utf8Length = 0;
    }
}
#endif
