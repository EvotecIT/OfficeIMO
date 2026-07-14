#nullable enable

using System.Buffers;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    internal sealed class CsvTextDataReaderRowSource : IDisposable
    {
        private readonly string _text;
        private readonly CsvLoadOptions _options;
        private CsvTextFieldSpanReadState _state;
        private CsvDataReaderTextRowVisitor _visitor;
        private bool _disposed;

        internal CsvTextDataReaderRowSource(
            string text,
            CsvLoadOptions options,
            int recordsToSkip,
            int sourceColumnCount)
        {
            _text = text;
            _options = options;
            _state = CreateTextFieldSpanReadState(text.AsSpan(), options, recordsToSkip);
            _visitor = new CsvDataReaderTextRowVisitor(text, sourceColumnCount);
        }

        public bool Read()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            _visitor.Reset();
            if (!TryReadNextTextRecordFieldSpans(_text.AsSpan(), _options, null, ref _state, ref _visitor, out var fieldCount))
            {
                return false;
            }

            _visitor.Complete(fieldCount, _options.ColumnCountMismatchPolicy);
            return true;
        }

        public ReadOnlySpan<char> GetSpan(int ordinal) => _visitor.GetSpan(ordinal);

        public string GetString(int ordinal) => _visitor.GetString(ordinal);

        public bool IsNull(int ordinal, string? nullValue)
        {
            return !_visitor.IsMissing(ordinal) &&
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

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
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
        private bool _nextVisitIsUnescapedScratch;

        internal CsvDataReaderTextRowVisitor(string text, int sourceColumnCount)
        {
            _text = text;
            _starts = new int[sourceColumnCount];
            _lengths = new int[sourceColumnCount];
            _materialized = new string?[sourceColumnCount];
            _nextVisitIsUnescapedScratch = false;
        }

        internal int SourceColumnCount => _starts.Length;

        internal void Reset()
        {
            _nextVisitIsUnescapedScratch = false;
        }

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            if ((uint)fieldIndex >= (uint)_starts.Length)
            {
                _nextVisitIsUnescapedScratch = false;
                return;
            }

            _lengths[fieldIndex] = value.Length;
            if (_nextVisitIsUnescapedScratch)
            {
                _starts[fieldIndex] = -1;
                _materialized[fieldIndex] = value.Length == 0 ? string.Empty : value.ToString();
                _nextVisitIsUnescapedScratch = false;
                return;
            }

            ref char textStart = ref MemoryMarshal.GetReference(_text.AsSpan());
            ref char valueStart = ref MemoryMarshal.GetReference(value);
            var byteOffset = Unsafe.ByteOffset(ref textStart, ref valueStart);
            _starts[fieldIndex] = checked((int)(byteOffset / 2));
            _materialized[fieldIndex] = null;
        }

        public bool TryVisitEscapedField(int recordIndex, int fieldIndex, ReadOnlySpan<char> escapedValue, int unescapedLength)
        {
            _nextVisitIsUnescapedScratch = true;
            return false;
        }

        public void VisitFieldValue(int recordIndex, int fieldIndex, string value)
        {
            _nextVisitIsUnescapedScratch = false;
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

    private static CsvTextFieldSpanReadState CreateTextFieldSpanReadState(
        ReadOnlySpan<char> text,
        CsvLoadOptions options,
        int recordsToSkip)
    {
        var delimiter = GetDelimiterChar(options);
        var delimiterVector = System.Runtime.Intrinsics.Vector256<byte>.Zero;
        if (!options.TrimWhitespace &&
            delimiter <= byte.MaxValue &&
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
