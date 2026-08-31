#nullable enable

using System.Buffers;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    internal sealed class CsvTextDataReaderRowSource : ICsvDataReaderTextRowSource
    {
        private readonly string _text;
        private readonly int _endPosition;
        private readonly CsvLoadOptions _options;
        private CsvTextFieldSpanReadState _state;
        private CsvDataReaderTextRowVisitor _visitor;
        private CsvTextDataReaderBatch? _batch;
        private bool _usingBatchRow;
        private bool _disposed;

        internal CsvTextDataReaderRowSource(
            string text,
            CsvLoadOptions options,
            int recordsToSkip,
            int sourceColumnCount)
            : this(text, options, recordsToSkip, sourceColumnCount, 0, text.Length)
        {
        }

        internal CsvTextDataReaderRowSource(
            string text,
            CsvLoadOptions options,
            int recordsToSkip,
            int sourceColumnCount,
            int startPosition,
            int endPosition)
        {
            if ((uint)startPosition > (uint)text.Length)
            {
                throw new ArgumentOutOfRangeException(nameof(startPosition));
            }
            if ((uint)endPosition > (uint)text.Length || endPosition < startPosition)
            {
                throw new ArgumentOutOfRangeException(nameof(endPosition));
            }
            _text = text;
            _endPosition = endPosition;
            _options = options;
            _state = CreateTextFieldSpanReadState(
                text.AsSpan(startPosition, endPosition - startPosition),
                options,
                recordsToSkip);
            _state.Position = startPosition;
            _visitor = new CsvDataReaderTextRowVisitor(text, sourceColumnCount);
            _batch = sourceColumnCount is > 0 and <= TextQuoteAwareFieldSpanCapacity
                ? new CsvTextDataReaderBatch(
                    text,
                    sourceColumnCount,
                    options.Culture,
                    options.DateTimeFormats)
                : null;
        }

        internal int PreferredParallelBatchSize => GetPreferredTextParallelBatchSize();

        internal bool CanTakeParallelBatch => !_disposed && _batch is not null && !_batch.HasStarted;

        internal int SourceColumnCount => _visitor.SourceColumnCount;

        internal CsvLoadOptions Options => _options;

        internal int PrepareForParallelPartition(CancellationToken cancellationToken)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            cancellationToken.ThrowIfCancellationRequested();
            SkipPendingTextDataReaderRecords(
                _text.AsSpan(0, _endPosition),
                _options,
                cancellationToken,
                ref _state);
            return _state.Position;
        }

        public bool Read()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_batch is not null)
            {
                if (_batch.MoveNext() ||
                    (TryFillTextDataReaderBatchAvx2(_text.AsSpan(0, _endPosition), _options, ref _state, _batch) && _batch.MoveNext()))
                {
                    _usingBatchRow = true;
                    ValidateFieldCount(_batch.CurrentFieldCount);
                    return true;
                }
            }

            _usingBatchRow = false;
            if (!TryReadNextTextRecordFieldSpans(_text.AsSpan(0, _endPosition), _options, null, ref _state, ref _visitor, out var fieldCount))
            {
                return false;
            }

            ValidateFieldCount(fieldCount);
            _visitor.Complete(fieldCount, _options.ColumnCountMismatchPolicy);
            return true;
        }

        internal bool TryTakeParallelBatch(
            int preferredBatchSize,
            CancellationToken cancellationToken,
            out ICsvDataReaderTextRowSource? rows)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            cancellationToken.ThrowIfCancellationRequested();
            rows = null;
            if (_batch is null || _batch.HasStarted)
            {
                return false;
            }

            int requestedRowCapacity = Math.Max(
                1,
                Math.Min(MaximumTextDataReaderBatchRowCapacity, preferredBatchSize));
            if (_batch.RowCapacity != requestedRowCapacity)
            {
                _batch.Dispose();
                _batch = new CsvTextDataReaderBatch(
                    _text,
                    _visitor.SourceColumnCount,
                    _options.Culture,
                    _options.DateTimeFormats,
                    requestedRowCapacity);
            }

            SkipPendingTextDataReaderRecords(_text.AsSpan(0, _endPosition), _options, cancellationToken, ref _state);

            if (_state.Position >= _endPosition)
            {
                return true;
            }

            if (!TryFillTextDataReaderBatchAvx2(
                    _text.AsSpan(0, _endPosition),
                    _options,
                    ref _state,
                    _batch,
                    cancellationToken) &&
                !TryFillTextDataReaderBatchScalar(
                    _text,
                    _options,
                    ref _state,
                    _batch,
                    cancellationToken))
            {
                return false;
            }

            CsvTextDataReaderBatch detached = _batch;
            _batch = new CsvTextDataReaderBatch(
                _text,
                detached.SourceColumnCount,
                _options.Culture,
                _options.DateTimeFormats,
                preferredBatchSize);
            rows = detached;
            return true;
        }

        public ReadOnlySpan<char> GetSpan(int ordinal) => _usingBatchRow
            ? _batch!.GetSpan(ordinal)
            : _visitor.GetSpan(ordinal);

        public string GetString(int ordinal) => _usingBatchRow
            ? _batch!.GetString(ordinal)
            : _visitor.GetString(ordinal);

        public bool IsMissing(int ordinal) => _usingBatchRow
            ? _batch!.IsMissing(ordinal)
            : _visitor.IsMissing(ordinal);

        public bool IsNull(int ordinal, string? nullValue)
        {
            var isMissing = _usingBatchRow
                ? _batch!.IsMissing(ordinal)
                : _visitor.IsMissing(ordinal);
            return !isMissing &&
                nullValue is not null &&
                GetSpan(ordinal).SequenceEqual(nullValue.AsSpan());
        }

        public int CopyStringValues(object[] values, int count, string? nullValue)
        {
            var valueCount = Math.Min(count, _visitor.SourceColumnCount);
            for (var i = 0; i < valueCount; i++)
            {
                values[i] = nullValue is not null && IsNull(i, nullValue) ? DBNull.Value : GetString(i);
            }

            for (var i = valueCount; i < count; i++)
            {
                values[i] = DBNull.Value;
            }

            return count;
        }

        private void ValidateFieldCount(int fieldCount)
        {
            if (_options.ColumnCountMismatchPolicy == CsvColumnCountMismatchPolicy.Strict &&
                fieldCount != _visitor.SourceColumnCount)
            {
                throw new CsvException($"Row contains {fieldCount} values but header defines {_visitor.SourceColumnCount} columns.");
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _batch?.Dispose();
            _batch = null;
            if (_state.Scratch is not null)
            {
                ArrayPool<char>.Shared.Return(_state.Scratch);
                _state.Scratch = null;
            }
        }
    }

    private struct CsvDataReaderTextRowVisitor : ICsvFieldSpanVisitor
    {
        private readonly string _text;
        private readonly int[] _starts;
        private readonly int[] _lengths;
        private readonly string?[] _materialized;

        internal CsvDataReaderTextRowVisitor(string text, int sourceColumnCount)
        {
            _text = text;
            _starts = new int[sourceColumnCount];
            _lengths = new int[sourceColumnCount];
            _materialized = new string?[sourceColumnCount];
        }

        internal int SourceColumnCount => _starts.Length;

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            if ((uint)fieldIndex >= (uint)_starts.Length)
            {
                return;
            }

            _lengths[fieldIndex] = value.Length;
            ref char textStart = ref MemoryMarshal.GetReference(_text.AsSpan());
            ref char valueStart = ref MemoryMarshal.GetReference(value);
            var byteOffset = Unsafe.ByteOffset(ref textStart, ref valueStart);
            _starts[fieldIndex] = checked((int)(byteOffset / 2));
            _materialized[fieldIndex] = null;
        }

        public void VisitFieldRange(
            int recordIndex,
            int fieldIndex,
            char[] buffer,
            int start,
            int length)
        {
            VisitFieldValue(
                recordIndex,
                fieldIndex,
                length == 0 ? string.Empty : new string(buffer, start, length));
        }

        public bool TryVisitEscapedField(int recordIndex, int fieldIndex, ReadOnlySpan<char> escapedValue, int unescapedLength)
        {
            if ((uint)fieldIndex >= (uint)_starts.Length)
            {
                return true;
            }

            ref char textStart = ref MemoryMarshal.GetReference(_text.AsSpan());
            ref char valueStart = ref MemoryMarshal.GetReference(escapedValue);
            var byteOffset = Unsafe.ByteOffset(ref textStart, ref valueStart);
            var sourceStart = checked((int)(byteOffset / 2));
            _starts[fieldIndex] = -1;
            _lengths[fieldIndex] = unescapedLength;
            _materialized[fieldIndex] = string.Create(
                unescapedLength,
                (Text: _text, Start: sourceStart, Length: escapedValue.Length),
                static (destination, state) =>
                {
                    var source = state.Text.AsSpan(state.Start, state.Length);
                    var sourceIndex = 0;
                    var destinationIndex = 0;
                    while (sourceIndex < source.Length)
                    {
                        var value = source[sourceIndex++];
                        destination[destinationIndex++] = value;
                        if (value == '"' && sourceIndex < source.Length && source[sourceIndex] == '"')
                        {
                            sourceIndex++;
                        }
                    }
                });
            return true;
        }

        public void VisitFieldValue(int recordIndex, int fieldIndex, string value)
        {
            if ((uint)fieldIndex >= (uint)_starts.Length)
            {
                return;
            }

            _starts[fieldIndex] = -1;
            _lengths[fieldIndex] = value.Length;
            _materialized[fieldIndex] = value;
        }

        internal void Complete(int fieldCount, CsvColumnCountMismatchPolicy mismatchPolicy)
        {
            if (mismatchPolicy == CsvColumnCountMismatchPolicy.Strict && fieldCount != _starts.Length)
            {
                throw new CsvException($"Row contains {fieldCount} values but header defines {_starts.Length} columns.");
            }

            for (var i = Math.Min(fieldCount, _starts.Length); i < _starts.Length; i++)
            {
                _lengths[i] = -1;
                _materialized[i] = null;
            }
        }

        internal ReadOnlySpan<char> GetSpan(int ordinal)
        {
            if ((uint)ordinal >= (uint)_starts.Length)
            {
                throw new IndexOutOfRangeException();
            }

            if (_lengths[ordinal] < 0)
            {
                return ReadOnlySpan<char>.Empty;
            }

            var materialized = _materialized[ordinal];
            return materialized is not null
                ? materialized.AsSpan()
                : _text.AsSpan(_starts[ordinal], _lengths[ordinal]);
        }

        internal bool IsMissing(int ordinal)
        {
            if ((uint)ordinal >= (uint)_starts.Length)
            {
                throw new IndexOutOfRangeException();
            }

            return _lengths[ordinal] < 0;
        }

        internal string GetString(int ordinal)
        {
            if ((uint)ordinal >= (uint)_starts.Length)
            {
                throw new IndexOutOfRangeException();
            }

            if (_lengths[ordinal] <= 0)
            {
                return string.Empty;
            }

            var materialized = _materialized[ordinal];
            if (materialized is not null)
            {
                return materialized;
            }

            materialized = _text.Substring(_starts[ordinal], _lengths[ordinal]);
            _materialized[ordinal] = materialized;
            return materialized;
        }
    }

    internal static bool CanReadDataReaderRowsFromText(string text, CsvLoadOptions options)
    {
        return !HasFieldLengthLimits(options) &&
            !UsesTextDelimiter(options) &&
            !(NeedsLogicalCommentSkipping(options) && HasPotentialTextCommentRecord(text, options.CommentCharacter)) &&
            options.ParseErrorAction != CsvParseErrorAction.SkipRow &&
            !options.NormalizeQuotes &&
            !options.InternStrings &&
            options.StaticColumns is null;
    }

    internal static bool TryGetFirstTextDataReaderRecordFieldCount(
        string text,
        CsvLoadOptions options,
        int recordsToSkip,
        out int fieldCount)
    {
        var state = CreateTextFieldSpanReadState(text.AsSpan(), options, recordsToSkip);
        var visitor = new CsvFieldCountOnlyVisitor();
        try
        {
            return TryReadNextTextRecordFieldSpans(
                text.AsSpan(),
                options,
                null,
                ref state,
                ref visitor,
                out fieldCount);
        }
        finally
        {
            if (state.Scratch is not null)
            {
                ArrayPool<char>.Shared.Return(state.Scratch);
            }
        }
    }

    private readonly struct CsvFieldCountOnlyVisitor : ICsvFieldSpanVisitor
    {
        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
        }

        public bool TryVisitEscapedField(
            int recordIndex,
            int fieldIndex,
            ReadOnlySpan<char> escapedValue,
            int unescapedLength) => true;
    }

    private static void SkipPendingTextDataReaderRecords(
        ReadOnlySpan<char> text,
        CsvLoadOptions options,
        CancellationToken cancellationToken,
        ref CsvTextFieldSpanReadState state)
    {
        if (state.RecordsToSkip == 0)
        {
            return;
        }

        var delimiter = GetDelimiterChar(options);
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var visitor = new CsvFieldCountOnlyVisitor();
        while (state.RecordsToSkip > 0 && state.Position < text.Length)
        {
            ThrowIfCancellationRequested(options);
            cancellationToken.ThrowIfCancellationRequested();
            int recordStart = state.Position;
            if (TrySkipTextEmptyRecord(text, trim, allowEmpty, ref state.Position))
            {
                continue;
            }

            bool startsWithCommentCharacter = text[state.Position] == options.CommentCharacter;
            bool isW3CFieldsHeader = startsWithCommentCharacter &&
                CanReadW3CFieldsHeader(options, state.EmittedRecordCount) &&
                IsTextW3CFieldsLine(text, state.Position);
            bool skipCommentRecord = startsWithCommentCharacter &&
                !isW3CFieldsHeader &&
                (options.SkipCommentRows ||
                    (options.HasHeaderRow &&
                        options.Header is null &&
                        options.SkipCommentRowsBeforeHeader &&
                        state.EmittedRecordCount <= GetParserInitialRecordsToSkip(options)));
            if (skipCommentRecord)
            {
                SkipTextRecord(text, ref state.Position);
                continue;
            }

            if (!trim &&
                TrySkipTextUnquotedRecord(text, delimiter, ref state.Position, out int skippedDelimiterCount))
            {
                state.UnquotedDelimiterIndexCapacity = GetTextDelimiterIndexCapacity(skippedDelimiterCount);
                state.RecordsToSkip--;
                continue;
            }

            int fieldCount;
            int firstFieldLength;
            try
            {
                if (!TryReadTextUnquotedRecordFieldSpans(
                        text,
                        delimiter,
                        trim,
                        allowEmpty,
                        emitFields: false,
                        state.RecordIndex,
                        ref state.UseAvx2UnquotedFastPath,
                        ref state.UnquotedDelimiterIndexCapacity,
                        state.TextMayContainQuote,
                        state.DelimiterVector,
                        ref state.Position,
                        projectedFieldVisitor: null,
                        ref visitor,
                        ref state.Scratch,
                        out fieldCount,
                        out firstFieldLength))
                {
                    fieldCount = ReadTextRecordFieldSpans(
                        text,
                        delimiter,
                        trim,
                        strictQuotes,
                        emitFields: false,
                        state.RecordIndex,
                        ref state.Position,
                        projectedFieldVisitor: null,
                        ref visitor,
                        ref state.Scratch,
                        out firstFieldLength);
                }
            }
            catch (CsvParseException ex) when (HandleParseError(options, ex, state.LineNumber))
            {
                state.Position = recordStart;
                SkipTextRecord(text, ref state.Position);
                state.LineNumber++;
                continue;
            }

            int requiredFieldCapacity = GetTextDelimiterIndexCapacity(fieldCount);
            if (requiredFieldCapacity > state.UnquotedDelimiterIndexCapacity)
            {
                state.UnquotedDelimiterIndexCapacity = requiredFieldCapacity;
            }

            bool isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
            if (fieldCount != 0 && (allowEmpty || !isEmptyRecord))
            {
                state.RecordsToSkip--;
            }

            if (state.Position == recordStart)
            {
                state.Position = text.Length;
            }
        }
    }

    private static CsvTextFieldSpanReadState CreateTextFieldSpanReadState(
        ReadOnlySpan<char> text,
        CsvLoadOptions options,
        int recordsToSkip)
    {
        var delimiter = GetDelimiterChar(options);
        var delimiterVector = System.Runtime.Intrinsics.Vector256<byte>.Zero;
        if (!options.TrimWhitespace &&
            CanUseAvx2PackedDelimiter(delimiter) &&
            System.Runtime.Intrinsics.X86.Avx2.IsSupported)
        {
            delimiterVector = System.Runtime.Intrinsics.Vector256.Create((byte)delimiter);
        }

        return new CsvTextFieldSpanReadState
        {
            RecordsToSkip = recordsToSkip,
            UseAvx2UnquotedFastPath = true,
            TextMayContainQuote = text.Length < TextQuoteFreeProbeMinimumLength || text.IndexOf('"') >= 0,
            UnquotedDelimiterIndexCapacity = 16,
            DelimiterVector = delimiterVector,
            LineNumber = 1
        };
    }

    private static bool TryReadNextTextRecordFieldSpans<TVisitor>(
        ReadOnlySpan<char> text,
        CsvLoadOptions options,
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
        ref CsvTextFieldSpanReadState state,
        ref TVisitor fieldVisitor,
        out int emittedFieldCount)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var delimiter = GetDelimiterChar(options);
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        emittedFieldCount = 0;

        while (state.Position < text.Length)
        {
            ThrowIfCancellationRequested(options);
            var recordStart = state.Position;
            if (TrySkipTextEmptyRecord(text, trim, allowEmpty, ref state.Position))
            {
                continue;
            }

            var startsWithCommentCharacter = text[state.Position] == options.CommentCharacter;
            var isW3CFieldsHeader = startsWithCommentCharacter &&
                CanReadW3CFieldsHeader(options, state.EmittedRecordCount) &&
                IsTextW3CFieldsLine(text, state.Position);
            var skipCommentRecord = startsWithCommentCharacter &&
                !isW3CFieldsHeader &&
                (options.SkipCommentRows ||
                    (options.HasHeaderRow &&
                        options.Header is null &&
                        options.SkipCommentRowsBeforeHeader &&
                        state.EmittedRecordCount <= GetParserInitialRecordsToSkip(options)));
            if (skipCommentRecord)
            {
                SkipTextRecord(text, ref state.Position);
                continue;
            }

            if (state.RecordsToSkip > 0 &&
                !trim &&
                TrySkipTextUnquotedRecord(text, delimiter, ref state.Position, out var skippedDelimiterCount))
            {
                state.UnquotedDelimiterIndexCapacity = GetTextDelimiterIndexCapacity(skippedDelimiterCount);
                state.RecordsToSkip--;
                continue;
            }

            var emitFields = state.RecordsToSkip == 0;
            int fieldCount;
            int firstFieldLength;
            try
            {
                if (!TryReadTextUnquotedRecordFieldSpans(
                        text,
                        delimiter,
                        trim,
                        allowEmpty,
                        emitFields,
                        state.RecordIndex,
                        ref state.UseAvx2UnquotedFastPath,
                        ref state.UnquotedDelimiterIndexCapacity,
                        state.TextMayContainQuote,
                        state.DelimiterVector,
                        ref state.Position,
                        projectedFieldVisitor,
                        ref fieldVisitor,
                        ref state.Scratch,
                        out fieldCount,
                        out firstFieldLength))
                {
                    fieldCount = ReadTextRecordFieldSpans(
                        text,
                        delimiter,
                        trim,
                        strictQuotes,
                        emitFields,
                        state.RecordIndex,
                        ref state.Position,
                        projectedFieldVisitor,
                        ref fieldVisitor,
                        ref state.Scratch,
                        out firstFieldLength);
                }
            }
            catch (CsvParseException ex) when (HandleParseError(options, ex, state.LineNumber))
            {
                state.Position = recordStart;
                SkipTextRecord(text, ref state.Position);
                state.LineNumber++;
                continue;
            }

            var requiredFieldCapacity = GetTextDelimiterIndexCapacity(fieldCount);
            if (requiredFieldCapacity > state.UnquotedDelimiterIndexCapacity)
            {
                state.UnquotedDelimiterIndexCapacity = requiredFieldCapacity;
            }

            var isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
            var shouldEmit = fieldCount != 0 && (allowEmpty || !isEmptyRecord);
            if (!shouldEmit)
            {
                continue;
            }

            if (state.RecordsToSkip > 0)
            {
                state.RecordsToSkip--;
                continue;
            }

            state.RecordIndex++;
            state.EmittedRecordCount++;
            ReportProgress(options, state.EmittedRecordCount, state.LineNumber);
            state.LineNumber++;
            emittedFieldCount = fieldCount;

            if (state.Position == recordStart)
            {
                state.Position = text.Length;
            }

            return true;
        }

        return false;
    }

    private struct CsvTextFieldSpanReadState
    {
        public int Position;
        public int RecordsToSkip;
        public int RecordIndex;
        public int EmittedRecordCount;
        public int LineNumber;
        public bool UseAvx2UnquotedFastPath;
        public bool TextMayContainQuote;
        public int UnquotedDelimiterIndexCapacity;
        public System.Runtime.Intrinsics.Vector256<byte> DelimiterVector;
        public char[]? Scratch;
    }
#endif
}
