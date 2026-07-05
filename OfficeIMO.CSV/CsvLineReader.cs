#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

internal sealed class CsvLineReader
{
    private const int DefaultBufferSize = 32 * 1024;
    private readonly TextReader _reader;
    private readonly char[] _buffer;
    private int _position;
    private int _length;
    private bool _endOfReader;

    public CsvLineReader(TextReader reader)
    {
        _reader = reader ?? throw new ArgumentNullException(nameof(reader));
        _buffer = new char[DefaultBufferSize];
    }

    public CsvLineReadResult ReadUnquotedRecordOrLine(char delimiter, bool trim, char commentCharacter, List<string> fields, out string? line, out string separator)
    {
        fields.Clear();
        line = null;
        separator = string.Empty;

        if (!EnsureBuffered())
        {
            return CsvLineReadResult.EndOfReader;
        }

        var segmentStart = _position;
        var segmentLength = _length - _position;
        if (segmentLength == 0)
        {
            return CsvLineReadResult.EndOfReader;
        }

        if (_buffer[segmentStart] == commentCharacter)
        {
            line = ReadLine(out separator);
            return CsvLineReadResult.Line;
        }

        var newlineIndex = IndexOfNewline(segmentStart, segmentLength);
        if (newlineIndex < 0)
        {
            line = ReadLine(out separator);
            return CsvLineReadResult.Line;
        }

        var lineLength = newlineIndex - segmentStart;
        if (Array.IndexOf(_buffer, '"', segmentStart, lineLength) >= 0)
        {
            line = ReadLine(out separator);
            return CsvLineReadResult.Line;
        }

        AddUnquotedFields(segmentStart, newlineIndex, delimiter, trim, fields);
        _position = newlineIndex;
        ConsumeLineSeparator(_buffer[newlineIndex], out separator);
        return CsvLineReadResult.UnquotedRecord;
    }

#if NET8_0_OR_GREATER
    public CsvLineReadResult ReadUnquotedFieldSpansOrLine<TVisitor>(
        char delimiter,
        bool trim,
        char commentCharacter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        ref TVisitor fieldVisitor,
        out int fieldCount,
        out bool isEmptyRecord,
        out string? line,
        out string separator)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        isEmptyRecord = false;
        line = null;
        separator = string.Empty;

        if (!EnsureBuffered())
        {
            return CsvLineReadResult.EndOfReader;
        }

        var segmentStart = _position;
        var segmentLength = _length - _position;
        if (segmentLength == 0)
        {
            return CsvLineReadResult.EndOfReader;
        }

        if (_buffer[segmentStart] == commentCharacter)
        {
            line = ReadLine(out separator);
            return CsvLineReadResult.Line;
        }

        if (!trim &&
            delimiter <= byte.MaxValue &&
            System.Runtime.Intrinsics.X86.Avx2.IsSupported &&
            TryReadUnquotedFieldSpansOrLineAvx2(
                delimiter,
                allowEmpty,
                emitFields,
                recordIndex,
                ref fieldVisitor,
                out fieldCount,
                out isEmptyRecord,
                out separator,
                out var readResult))
        {
            return readResult;
        }

        var specialIndex = _buffer.AsSpan(segmentStart, segmentLength).IndexOfAny('"', '\r', '\n');
        while (specialIndex < 0 && TryExtendCurrentSegment())
        {
            segmentStart = _position;
            segmentLength = _length - _position;
            specialIndex = _buffer.AsSpan(segmentStart, segmentLength).IndexOfAny('"', '\r', '\n');
        }

        var endsAtReaderEnd = false;
        if (specialIndex < 0)
        {
            if (!_endOfReader || segmentStart == 0 && segmentLength == _buffer.Length)
            {
                line = ReadLine(out separator);
                return CsvLineReadResult.Line;
            }

            specialIndex = segmentLength;
            endsAtReaderEnd = true;
        }

        var newlineIndex = segmentStart + specialIndex;
        if (!endsAtReaderEnd && _buffer[newlineIndex] == '"')
        {
            line = ReadLine(out separator);
            return CsvLineReadResult.Line;
        }

        var lineLength = newlineIndex - segmentStart;
        var emitNonEmptyRecord = allowEmpty || (trim
            ? HasNonWhitespaceOrDelimiter(segmentStart, newlineIndex, delimiter)
            : lineLength != 0);
        fieldCount = VisitUnquotedFieldSpans(segmentStart, newlineIndex, delimiter, trim, emitNonEmptyRecord, emitFields, recordIndex, ref fieldVisitor, out var firstFieldLength);
        isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
        _position = newlineIndex;
        if (endsAtReaderEnd)
        {
            separator = string.Empty;
        }
        else
        {
            ConsumeLineSeparator(_buffer[newlineIndex], out separator);
        }

        return CsvLineReadResult.UnquotedRecord;
    }
#endif

    public string? ReadLine(out string separator)
    {
        separator = string.Empty;
        StringBuilder? builder = null;

        while (true)
        {
            if (!EnsureBuffered())
            {
                return builder?.ToString();
            }

            var segmentStart = _position;
            var newlineIndex = IndexOfNewline(_position, _length - _position);
            if (newlineIndex >= 0)
            {
                _position = newlineIndex;
                return CompleteLine(builder, segmentStart, newlineIndex, _buffer[newlineIndex], out separator);
            }

            _position = _length;
            builder ??= new StringBuilder(Math.Max(256, _length - segmentStart));
            builder.Append(_buffer, segmentStart, _length - segmentStart);
        }
    }

    private int IndexOfNewline(int start, int count)
    {
        var lineFeed = Array.IndexOf(_buffer, '\n', start, count);
        var carriageReturn = Array.IndexOf(_buffer, '\r', start, count);
        if (lineFeed < 0)
        {
            return carriageReturn;
        }

        if (carriageReturn < 0)
        {
            return lineFeed;
        }

        return Math.Min(lineFeed, carriageReturn);
    }

    private void AddUnquotedFields(int start, int end, char delimiter, bool trim, List<string> fields)
    {
        var fieldStart = start;
        for (var i = start; i < end; i++)
        {
            if (_buffer[i] != delimiter)
            {
                continue;
            }

            fields.Add(GetUnquotedField(fieldStart, i - fieldStart, trim));
            fieldStart = i + 1;
        }

        fields.Add(GetUnquotedField(fieldStart, end - fieldStart, trim));
    }

    private string GetUnquotedField(int start, int length, bool trim)
    {
        if (length == 0)
        {
            return string.Empty;
        }

        if (!trim)
        {
            return new string(_buffer, start, length);
        }

        var end = start + length - 1;
        while (start <= end && char.IsWhiteSpace(_buffer[start]))
        {
            start++;
        }

        while (end >= start && char.IsWhiteSpace(_buffer[end]))
        {
            end--;
        }

        return end < start ? string.Empty : new string(_buffer, start, end - start + 1);
    }

#if NET8_0_OR_GREATER
    private bool HasNonWhitespaceOrDelimiter(int start, int end, char delimiter)
    {
        for (var i = start; i < end; i++)
        {
            var value = _buffer[i];
            if (value == delimiter || !char.IsWhiteSpace(value))
            {
                return true;
            }
        }

        return false;
    }

    private int VisitUnquotedFieldSpans<TVisitor>(
        int start,
        int end,
        char delimiter,
        bool trim,
        bool emitNonEmptyRecord,
        bool emitFields,
        int recordIndex,
        ref TVisitor fieldVisitor,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (!trim)
        {
            return VisitUntrimmedUnquotedFieldSpans(start, end, delimiter, emitNonEmptyRecord && emitFields, recordIndex, ref fieldVisitor, out firstFieldLength);
        }

        var fieldIndex = 0;
        var fieldStart = start;
        var emit = emitNonEmptyRecord && emitFields;
        firstFieldLength = 0;
        for (var i = start; i < end; i++)
        {
            if (_buffer[i] != delimiter)
            {
                continue;
            }

            VisitFieldSpan(fieldStart, i - fieldStart, trim: true, emit, recordIndex, fieldIndex, ref fieldVisitor, ref firstFieldLength);
            fieldIndex++;
            fieldStart = i + 1;
        }

        VisitFieldSpan(fieldStart, end - fieldStart, trim: true, emit, recordIndex, fieldIndex, ref fieldVisitor, ref firstFieldLength);
        return fieldIndex + 1;
    }

    private int VisitUntrimmedUnquotedFieldSpans<TVisitor>(
        int start,
        int end,
        char delimiter,
        bool emit,
        int recordIndex,
        ref TVisitor fieldVisitor,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var fieldIndex = 0;
        var fieldStart = start;
        firstFieldLength = 0;
        for (var i = start; i < end; i++)
        {
            if (_buffer[i] != delimiter)
            {
                continue;
            }

            var length = i - fieldStart;
            if (fieldIndex == 0)
            {
                firstFieldLength = length;
            }

            if (emit)
            {
                fieldVisitor.VisitField(recordIndex, fieldIndex, _buffer.AsSpan(fieldStart, length));
            }

            fieldIndex++;
            fieldStart = i + 1;
        }

        var finalLength = end - fieldStart;
        if (fieldIndex == 0)
        {
            firstFieldLength = finalLength;
        }

        if (emit)
        {
            fieldVisitor.VisitField(recordIndex, fieldIndex, _buffer.AsSpan(fieldStart, finalLength));
        }

        return fieldIndex + 1;
    }

    private int VisitIndexedUntrimmedUnquotedFieldSpans<TVisitor>(
        int start,
        int end,
        ReadOnlySpan<int> delimiterIndexes,
        bool emit,
        int recordIndex,
        ref TVisitor fieldVisitor,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var fieldIndex = 0;
        var fieldStart = start;
        firstFieldLength = 0;
        foreach (var delimiterIndex in delimiterIndexes)
        {
            var length = delimiterIndex - fieldStart;
            if (fieldIndex == 0)
            {
                firstFieldLength = length;
            }

            if (emit)
            {
                fieldVisitor.VisitField(recordIndex, fieldIndex, _buffer.AsSpan(fieldStart, length));
            }

            fieldIndex++;
            fieldStart = delimiterIndex + 1;
        }

        var finalLength = end - fieldStart;
        if (fieldIndex == 0)
        {
            firstFieldLength = finalLength;
        }

        if (emit)
        {
            fieldVisitor.VisitField(recordIndex, fieldIndex, _buffer.AsSpan(fieldStart, finalLength));
        }

        return fieldIndex + 1;
    }

    private bool TryReadUnquotedFieldSpansOrLineAvx2<TVisitor>(
        char delimiter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        ref TVisitor fieldVisitor,
        out int fieldCount,
        out bool isEmptyRecord,
        out string separator,
        out CsvLineReadResult readResult)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        isEmptyRecord = false;
        separator = string.Empty;
        readResult = CsvLineReadResult.Line;

        var start = _position;
        var end = _length - 32;
        if (start > end)
        {
            return false;
        }

        Span<int> delimiterIndexes = stackalloc int[256];
        var delimiterCount = 0;
        var pos = start;
        var delimiterVector = System.Runtime.Intrinsics.Vector256.Create((byte)delimiter);
        var quoteVector = System.Runtime.Intrinsics.Vector256.Create((byte)'"');
        var carriageReturnVector = System.Runtime.Intrinsics.Vector256.Create((byte)'\r');
        var lineFeedVector = System.Runtime.Intrinsics.Vector256.Create((byte)'\n');

        while (pos <= end)
        {
            var values = System.Runtime.InteropServices.MemoryMarshal.Cast<char, short>(_buffer.AsSpan(pos, 32));
            var first = System.Runtime.Intrinsics.Vector256.LoadUnsafe(ref System.Runtime.InteropServices.MemoryMarshal.GetReference(values));
            var second = System.Runtime.Intrinsics.Vector256.LoadUnsafe(ref System.Runtime.InteropServices.MemoryMarshal.GetReference(values.Slice(16)));
            var packed = System.Runtime.Intrinsics.X86.Avx2.PackUnsignedSaturate(first, second);
            var packedBytes = System.Runtime.Intrinsics.Vector256.AsByte(
                System.Runtime.Intrinsics.X86.Avx2.Permute4x64(System.Runtime.Intrinsics.Vector256.AsInt64(packed), 0b11_01_10_00));

            var delimiterMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, delimiterVector));
            var quoteMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, quoteVector));
            var carriageReturnMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, carriageReturnVector));
            var lineFeedMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, lineFeedVector));
            var terminalMask = quoteMask | carriageReturnMask | lineFeedMask;

            if (terminalMask != 0)
            {
                var terminalOffset = System.Numerics.BitOperations.TrailingZeroCount(terminalMask);
                var delimiterMaskBeforeTerminal = delimiterMask & ((1u << terminalOffset) - 1u);
                if (!AddDelimiterIndexes(delimiterMaskBeforeTerminal, pos, delimiterIndexes, ref delimiterCount))
                {
                    return false;
                }

                if (((quoteMask >> terminalOffset) & 1u) != 0)
                {
                    return false;
                }

                var newlineIndex = pos + terminalOffset;
                var lineLength = newlineIndex - start;
                fieldCount = VisitIndexedUntrimmedUnquotedFieldSpans(
                    start,
                    newlineIndex,
                    delimiterIndexes.Slice(0, delimiterCount),
                    (allowEmpty || lineLength != 0) && emitFields,
                    recordIndex,
                    ref fieldVisitor,
                    out var firstFieldLength);
                isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
                _position = newlineIndex;
                ConsumeLineSeparator(_buffer[newlineIndex], out separator);
                readResult = CsvLineReadResult.UnquotedRecord;
                return true;
            }

            if (!AddDelimiterIndexes(delimiterMask, pos, delimiterIndexes, ref delimiterCount))
            {
                return false;
            }

            pos += 32;
        }

        return false;
    }

    private static bool AddDelimiterIndexes(uint delimiterMask, int chunkStart, Span<int> delimiterIndexes, ref int delimiterCount)
    {
        while (delimiterMask != 0)
        {
            if (delimiterCount == delimiterIndexes.Length)
            {
                return false;
            }

            var offset = System.Numerics.BitOperations.TrailingZeroCount(delimiterMask);
            delimiterIndexes[delimiterCount++] = chunkStart + offset;
            delimiterMask &= delimiterMask - 1;
        }

        return true;
    }

    private void VisitFieldSpan<TVisitor>(
        int start,
        int length,
        bool trim,
        bool emit,
        int recordIndex,
        int fieldIndex,
        ref TVisitor fieldVisitor,
        ref int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var span = _buffer.AsSpan(start, length);
        if (trim)
        {
            span = span.Trim();
        }

        if (fieldIndex == 0)
        {
            firstFieldLength = span.Length;
        }

        if (emit)
        {
            fieldVisitor.VisitField(recordIndex, fieldIndex, span);
        }
    }
#endif

    private string CompleteLine(StringBuilder? builder, int segmentStart, int lineEnd, char newline, out string separator)
    {
        var line = builder is null
            ? new string(_buffer, segmentStart, lineEnd - segmentStart)
            : AppendAndGetString(builder, segmentStart, lineEnd - segmentStart);

        ConsumeLineSeparator(newline, out separator);
        return line;
    }

    private void ConsumeLineSeparator(char newline, out string separator)
    {
        _position++;
        if (newline == '\r')
        {
            separator = ConsumeLineFeedAfterCarriageReturn() ? "\r\n" : "\r";
        }
        else
        {
            separator = "\n";
        }
    }

    private string AppendAndGetString(StringBuilder builder, int start, int count)
    {
        if (count > 0)
        {
            builder.Append(_buffer, start, count);
        }

        return builder.ToString();
    }

    private bool ConsumeLineFeedAfterCarriageReturn()
    {
        if (!EnsureBuffered())
        {
            return false;
        }

        if (_buffer[_position] != '\n')
        {
            return false;
        }

        _position++;
        return true;
    }

    private bool EnsureBuffered()
    {
        if (_position < _length)
        {
            return true;
        }

        if (_endOfReader)
        {
            return false;
        }

        _length = _reader.Read(_buffer, 0, _buffer.Length);
        _position = 0;
        if (_length > 0)
        {
            return true;
        }

        _endOfReader = true;
        return false;
    }

#if NET8_0_OR_GREATER
    private bool TryExtendCurrentSegment()
    {
        var remaining = _length - _position;
        if (_position == 0 || remaining >= _buffer.Length)
        {
            return false;
        }

        if (remaining > 0)
        {
            Array.Copy(_buffer, _position, _buffer, 0, remaining);
        }

        _position = 0;
        _length = remaining;
        var read = _reader.Read(_buffer, remaining, _buffer.Length - remaining);
        if (read == 0)
        {
            _endOfReader = true;
            return false;
        }

        _length += read;
        return true;
    }
#endif
}

internal enum CsvLineReadResult
{
    EndOfReader,
    UnquotedRecord,
    Line
}
