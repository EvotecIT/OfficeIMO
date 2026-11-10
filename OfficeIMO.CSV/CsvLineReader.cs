using System.Buffers;
using System.Text;

namespace OfficeIMO.CSV;

internal sealed partial class CsvLineReader : IDisposable
{
    /// <summary>Default buffer size for streaming CSV text transport and record parsing.</summary>
    internal const int DefaultBufferSize = 32 * 1024;
    private const int MaximumReadRequestSize = 256 * 1024;
#if NET8_0_OR_GREATER
    private const int UnquotedDelimiterIndexCapacity = 64;
    private const int QuotedPrefixReuseMinimumDelimiterCount = 4;
    private static readonly System.Runtime.Intrinsics.Vector256<byte> CsvCommaVector =
        System.Runtime.Intrinsics.Vector256.Create((byte)',');
    private static readonly System.Runtime.Intrinsics.Vector256<byte> CsvQuoteVector =
        System.Runtime.Intrinsics.Vector256.Create((byte)'"');
    private static readonly System.Runtime.Intrinsics.Vector256<byte> CsvCarriageReturnVector =
        System.Runtime.Intrinsics.Vector256.Create((byte)'\r');
    private static readonly System.Runtime.Intrinsics.Vector256<byte> CsvLineFeedVector =
        System.Runtime.Intrinsics.Vector256.Create((byte)'\n');
#endif
    private readonly TextReader _reader;
    private char[] _buffer;
    private readonly CancellationToken _cancellationToken;
    private readonly CsvParsingSettings? _options;
    private CancellationToken _operationCancellationToken;
    private int _position;
    private int _length;
    private bool _endOfReader;
    private bool _skipLineFeedAfterCarriageReturn;

    internal char[] Buffer => _buffer;

    internal int PhysicalLineSeparatorsConsumed { get; private set; }

    internal bool TryGrowFilledBuffer(int minimumCapacity)
    {
        if (minimumCapacity <= _buffer.Length || _length < _buffer.Length)
        {
            return false;
        }

        char[] replacement = ArrayPool<char>.Shared.Rent(minimumCapacity);
        int remaining = _length - _position;
        if (remaining > 0)
        {
            Array.Copy(_buffer, _position, replacement, 0, remaining);
        }

        char[] previous = _buffer;
        _buffer = replacement;
        _position = 0;
        _length = remaining;
        ArrayPool<char>.Shared.Return(previous);
        return true;
    }

    public CsvLineReader(TextReader reader, CancellationToken cancellationToken)
    {
        _reader = reader ?? throw new ArgumentNullException(nameof(reader));
        _cancellationToken = cancellationToken;
        _buffer = ArrayPool<char>.Shared.Rent(DefaultBufferSize);
    }

    public CsvLineReader(TextReader reader, CsvParsingSettings options)
        : this(reader, options?.CancellationToken ?? throw new ArgumentNullException(nameof(options)))
    {
        _options = options;
    }

    internal void SetOperationCancellationToken(CancellationToken cancellationToken) =>
        _operationCancellationToken = cancellationToken;

    internal void ClearOperationCancellationToken() =>
        _operationCancellationToken = default;

    public void Dispose()
    {
        ArrayPool<char>.Shared.Return(_buffer);
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

#if NET8_0_OR_GREATER
        if (!trim &&
            CanUseAvx2PackedDelimiter(delimiter) &&
            (System.Runtime.Intrinsics.X86.Avx2.IsSupported || System.Runtime.Intrinsics.Arm.AdvSimd.IsSupported) &&
            TryReadUnquotedRecordOrLineSimd(delimiter, fields, out line, out separator, out var readResult))
        {
            return readResult;
        }
#endif

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
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
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
        LastQuotedRecordLineSeparators = -1;

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
            delimiter < byte.MaxValue &&
            System.Runtime.Intrinsics.X86.Avx512BW.IsSupported &&
            TryReadUnquotedFieldSpansOrLineAvx512(
                delimiter,
                allowEmpty,
                emitFields,
                recordIndex,
                projectedFieldVisitor,
                ref fieldVisitor,
                out fieldCount,
                out isEmptyRecord,
                out separator,
                out var avx512ReadResult))
        {
            return avx512ReadResult;
        }

        if (!trim &&
            CanUseAvx2PackedDelimiter(delimiter) &&
            (System.Runtime.Intrinsics.X86.Avx2.IsSupported || System.Runtime.Intrinsics.Arm.AdvSimd.IsSupported) &&
            TryReadUnquotedFieldSpansOrLineSimd(
                delimiter,
                allowEmpty,
                emitFields,
                recordIndex,
                projectedFieldVisitor,
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

        // A failed extension can still move the segment to the buffer start before it finds the end
        // of input, so the final record is read from the current position.
        segmentStart = _position;
        segmentLength = _length - _position;

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
            if (TryReadStandardQuotedFieldSpansOrLine(
                    delimiter,
                    trim,
                    allowEmpty,
                    emitFields,
                    recordIndex,
                    projectedFieldVisitor,
                    ref fieldVisitor,
                    out fieldCount,
                    out isEmptyRecord,
                    out separator,
                    out var quotedReadResult))
            {
                return quotedReadResult;
            }

            line = ReadLine(out separator);
            return CsvLineReadResult.Line;
        }

        var lineLength = newlineIndex - segmentStart;
        var emitNonEmptyRecord = allowEmpty || (trim
            ? HasNonWhitespaceOrDelimiter(segmentStart, newlineIndex, delimiter)
            : lineLength != 0);
        fieldCount = VisitUnquotedFieldSpans(segmentStart, newlineIndex, delimiter, trim, emitNonEmptyRecord, emitFields, recordIndex, projectedFieldVisitor, ref fieldVisitor, out var firstFieldLength);
        isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
        _position = newlineIndex;
        if (endsAtReaderEnd)
        {
            separator = string.Empty;
        }
        else
        {
            ConsumeLineSeparator(_buffer[newlineIndex], out separator, deferLineFeedRead: true);
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

    /// <summary>
    /// Finds the first CR or LF in one pass. Searching for each separator over the whole buffered
    /// text made every line of a file without the other separator scan the rest of the buffer.
    /// </summary>
    private int IndexOfNewline(int start, int count)
    {
#if NET8_0_OR_GREATER
        int index = _buffer.AsSpan(start, count).IndexOfAny('\r', '\n');
        return index < 0 ? -1 : start + index;
#else
        char[] buffer = _buffer;
        int end = start + count;
        for (int index = start; index < end; index++)
        {
            char value = buffer[index];
            if (value == '\n' || value == '\r')
            {
                return index;
            }
        }

        return -1;
#endif
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
    private void AddIndexedUnquotedFields(int start, int end, ReadOnlySpan<int> delimiterIndexes, List<string> fields)
    {
        var fieldStart = start;
        foreach (var delimiterIndex in delimiterIndexes)
        {
            fields.Add(GetUnquotedField(fieldStart, delimiterIndex - fieldStart, trim: false));
            fieldStart = delimiterIndex + 1;
        }

        fields.Add(GetUnquotedField(fieldStart, end - fieldStart, trim: false));
    }

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
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
        ref TVisitor fieldVisitor,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (!trim)
        {
            return VisitUntrimmedUnquotedFieldSpans(start, end, delimiter, emitNonEmptyRecord && emitFields, recordIndex, projectedFieldVisitor, ref fieldVisitor, out firstFieldLength);
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

            VisitFieldSpan(fieldStart, i - fieldStart, trim: true, emit, recordIndex, fieldIndex, projectedFieldVisitor, ref fieldVisitor, ref firstFieldLength);
            fieldIndex++;
            fieldStart = i + 1;
        }

        VisitFieldSpan(fieldStart, end - fieldStart, trim: true, emit, recordIndex, fieldIndex, projectedFieldVisitor, ref fieldVisitor, ref firstFieldLength);
        return fieldIndex + 1;
    }

    private int VisitUntrimmedUnquotedFieldSpans<TVisitor>(
        int start,
        int end,
        char delimiter,
        bool emit,
        int recordIndex,
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
        ref TVisitor fieldVisitor,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var fieldIndex = 0;
        var fieldStart = start;
        firstFieldLength = 0;
        if (projectedFieldVisitor is null)
        {
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
                    VisitFieldRange(ref fieldVisitor, recordIndex, fieldIndex, _buffer, fieldStart, length);
                }

                fieldIndex++;
                fieldStart = i + 1;
            }

            var unprojectedFinalLength = end - fieldStart;
            if (fieldIndex == 0)
            {
                firstFieldLength = unprojectedFinalLength;
            }

            if (emit)
            {
                VisitFieldRange(
                    ref fieldVisitor,
                    recordIndex,
                    fieldIndex,
                    _buffer,
                    fieldStart,
                    unprojectedFinalLength);
            }

            return fieldIndex + 1;
        }

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

            if (emit && CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
            {
                VisitFieldRange(ref fieldVisitor, recordIndex, fieldIndex, _buffer, fieldStart, length);
            }

            fieldIndex++;
            fieldStart = i + 1;
        }

        var finalLength = end - fieldStart;
        if (fieldIndex == 0)
        {
            firstFieldLength = finalLength;
        }

        if (emit && CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
        {
            VisitFieldRange(ref fieldVisitor, recordIndex, fieldIndex, _buffer, fieldStart, finalLength);
        }

        return fieldIndex + 1;
    }

    private int VisitIndexedUntrimmedUnquotedFieldSpans<TVisitor>(
        int start,
        int end,
        ReadOnlySpan<int> delimiterIndexes,
        bool emit,
        int recordIndex,
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
        ref TVisitor fieldVisitor,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var fieldIndex = 0;
        var fieldStart = start;
        firstFieldLength = 0;
        if (emit
            && projectedFieldVisitor is null
            && typeof(TVisitor) == typeof(CsvParser.CsvDataReaderStreamRowVisitor))
        {
            ref CsvParser.CsvDataReaderStreamRowVisitor streamVisitor = ref
                System.Runtime.CompilerServices.Unsafe.As<TVisitor, CsvParser.CsvDataReaderStreamRowVisitor>(
                    ref fieldVisitor);
            streamVisitor.VisitFieldRanges(_buffer, start, end, delimiterIndexes);
            firstFieldLength = delimiterIndexes.Length == 0
                ? end - start
                : delimiterIndexes[0] - start;
            return delimiterIndexes.Length + 1;
        }

        if (projectedFieldVisitor is null)
        {
            foreach (var delimiterIndex in delimiterIndexes)
            {
                var length = delimiterIndex - fieldStart;
                if (fieldIndex == 0)
                {
                    firstFieldLength = length;
                }

                if (emit)
                {
                    VisitFieldRange(ref fieldVisitor, recordIndex, fieldIndex, _buffer, fieldStart, length);
                }

                fieldIndex++;
                fieldStart = delimiterIndex + 1;
            }

            var unprojectedFinalLength = end - fieldStart;
            if (fieldIndex == 0)
            {
                firstFieldLength = unprojectedFinalLength;
            }

            if (emit)
            {
                VisitFieldRange(
                    ref fieldVisitor,
                    recordIndex,
                    fieldIndex,
                    _buffer,
                    fieldStart,
                    unprojectedFinalLength);
            }

            return fieldIndex + 1;
        }

        foreach (var delimiterIndex in delimiterIndexes)
        {
            var length = delimiterIndex - fieldStart;
            if (fieldIndex == 0)
            {
                firstFieldLength = length;
            }

            if (emit && CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
            {
                VisitFieldRange(ref fieldVisitor, recordIndex, fieldIndex, _buffer, fieldStart, length);
            }

            fieldIndex++;
            fieldStart = delimiterIndex + 1;
        }

        var finalLength = end - fieldStart;
        if (fieldIndex == 0)
        {
            firstFieldLength = finalLength;
        }

        if (emit && CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
        {
            VisitFieldRange(ref fieldVisitor, recordIndex, fieldIndex, _buffer, fieldStart, finalLength);
        }

        return fieldIndex + 1;
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static void VisitFieldRange<TVisitor>(
        ref TVisitor fieldVisitor,
        int recordIndex,
        int fieldIndex,
        char[] buffer,
        int start,
        int length)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (typeof(TVisitor) == typeof(CsvParser.CsvDataReaderStreamRowVisitor))
        {
            ref CsvParser.CsvDataReaderStreamRowVisitor streamVisitor = ref
                System.Runtime.CompilerServices.Unsafe.As<TVisitor, CsvParser.CsvDataReaderStreamRowVisitor>(
                    ref fieldVisitor);
            streamVisitor.VisitFieldRange(recordIndex, fieldIndex, buffer, start, length);
            return;
        }

        fieldVisitor.VisitField(recordIndex, fieldIndex, buffer.AsSpan(start, length));
    }


    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
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

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static bool CanUseAvx2PackedDelimiter(char delimiter) =>
        delimiter > byte.MinValue && delimiter < byte.MaxValue;

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static System.Runtime.Intrinsics.Vector256<byte> CreateDelimiterVector(char delimiter) =>
        delimiter == ','
            ? CsvCommaVector
            : System.Runtime.Intrinsics.Vector256.Create((byte)delimiter);

    private void VisitFieldSpan<TVisitor>(
        int start,
        int length,
        bool trim,
        bool emit,
        int recordIndex,
        int fieldIndex,
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
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

        if (emit && (projectedFieldVisitor is null || projectedFieldVisitor.ShouldVisitField(recordIndex, fieldIndex)))
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

    private void ConsumeLineSeparator(char newline, out string separator, bool deferLineFeedRead = false)
    {
        _position++;
        PhysicalLineSeparatorsConsumed++;
        if (newline == '\r')
        {
            if (deferLineFeedRead && _position >= _length && !_endOfReader)
            {
                _skipLineFeedAfterCarriageReturn = true;
                separator = "\r";
                return;
            }

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
        ThrowIfCancellationRequested();
        while (true)
        {
            if (_position < _length)
            {
                if (_skipLineFeedAfterCarriageReturn)
                {
                    _skipLineFeedAfterCarriageReturn = false;
                    if (_buffer[_position] == '\n')
                    {
                        _position++;
                        if (_position >= _length)
                        {
                            continue;
                        }
                    }
                }

                return true;
            }

            if (_endOfReader)
            {
                _skipLineFeedAfterCarriageReturn = false;
                return false;
            }

            _length = _reader.Read(_buffer, 0, Math.Min(MaximumReadRequestSize, _buffer.Length));
            ThrowIfCancellationRequested();
            _position = 0;
            if (_length == 0)
            {
                _endOfReader = true;
            }
        }
    }

#if NET8_0_OR_GREATER
    private bool TryExtendCurrentSegment()
    {
        ThrowIfCancellationRequested();
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
        var read = _reader.Read(
            _buffer,
            remaining,
            Math.Min(MaximumReadRequestSize, _buffer.Length - remaining));
        ThrowIfCancellationRequested();
        if (read == 0)
        {
            _endOfReader = true;
            return false;
        }

        _length += read;
        return true;
    }

#endif

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private void ThrowIfCancellationRequested()
    {
        _operationCancellationToken.ThrowIfCancellationRequested();
        _options?.OperationCancellationToken.ThrowIfCancellationRequested();
        _cancellationToken.ThrowIfCancellationRequested();
    }
}

internal enum CsvLineReadResult
{
    EndOfReader,
    UnquotedRecord,
    Line
}
