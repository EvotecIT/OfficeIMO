#nullable enable

#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    /// <summary>
    /// Pull-based row source over the existing bounded streaming parser. Field spans point at
    /// the reusable line buffer and are materialized only when a caller requests a string.
    /// </summary>
    internal sealed class CsvStreamDataReaderRowSource : ICsvDataReaderHeaderRowSource, ICsvDataReaderPositionSource
    {
        private const int LargeDataReaderBufferSize = 128 * 1024;
        private readonly TextReader _reader;
        private readonly CsvLineReader _lineReader;
        private readonly CsvLoadOptions _options;
        private readonly char _delimiter;
        private readonly bool _trim;
        private readonly bool _strictQuotes;
        private readonly bool _allowEmpty;
        private readonly Queue<CsvLine> _pendingLines = new();
        private readonly List<string> _quotedFields = new(32);
        private CsvDataReaderStreamRowVisitor _visitor;
        private readonly int _physicalLineOffset;
        private int _lineNumber;
        private int _emittedRecordCount;
        private int? _currentPhysicalLineNumber;
        private int? _currentPhysicalEndLineNumber;
        private bool _disposed;

        internal CsvStreamDataReaderRowSource(
            TextReader reader,
            CsvLoadOptions options,
            int initialEmittedRecordCount = 0,
            int physicalLineOffset = 0)
        {
            _reader = reader ?? throw new ArgumentNullException(nameof(reader));
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _delimiter = GetDelimiterChar(options);
            _trim = options.TrimWhitespace;
            _strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
            _allowEmpty = options.AllowEmptyLines;
            _lineReader = new CsvLineReader(reader, options.CancellationToken);
            _visitor = new CsvDataReaderStreamRowVisitor(_lineReader.Buffer);
            _emittedRecordCount = initialEmittedRecordCount;
            _physicalLineOffset = physicalLineOffset;
            _lineNumber = physicalLineOffset + 1;
        }

        internal int FieldCount => _visitor.FieldCount;

        int ICsvDataReaderHeaderRowSource.FieldCount => _visitor.FieldCount;

        int? ICsvDataReaderPositionSource.CurrentPhysicalLineNumber => _currentPhysicalLineNumber;

        int? ICsvDataReaderPositionSource.CurrentPhysicalEndLineNumber => _currentPhysicalEndLineNumber;

        internal void SetSourceColumnCount(int sourceColumnCount)
        {
            if (_lineReader.TryGrowFilledBuffer(LargeDataReaderBufferSize))
            {
                _visitor.SetBuffer(_lineReader.Buffer);
            }

            _visitor.SetSourceColumnCount(sourceColumnCount);
        }

        void ICsvDataReaderHeaderRowSource.SetSourceColumnCount(int sourceColumnCount) =>
            SetSourceColumnCount(sourceColumnCount);

        public bool Read()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            _visitor.Reset();
            _currentPhysicalLineNumber = null;
            _currentPhysicalEndLineNumber = null;
            while (true)
            {
                ThrowIfCancellationRequested(_options);
                bool recordStartedFromPendingLine = _pendingLines.Count > 0;
                int recordStartLineNumber = recordStartedFromPendingLine
                    ? _pendingLines.Peek().PhysicalLineNumber
                    : _physicalLineOffset + _lineReader.PhysicalLineSeparatorsConsumed + 1;
                string? fastLine = null;
                string lineSeparator;
                CsvLineReadResult readResult;
                if (_pendingLines.Count == 0)
                {
                    readResult = _lineReader.ReadUnquotedFieldSpansOrLine(
                        _delimiter,
                        _trim,
                        _options.CommentCharacter,
                        _allowEmpty,
                        emitFields: true,
                        recordIndex: _emittedRecordCount,
                        projectedFieldVisitor: null,
                        ref _visitor,
                        out int fieldCount,
                        out bool isEmptyRecord,
                        out fastLine,
                        out lineSeparator);

                    if (readResult == CsvLineReadResult.EndOfReader)
                    {
                        return false;
                    }

                    if (readResult == CsvLineReadResult.UnquotedRecord)
                    {
                        _lineNumber++;
                        if (fieldCount == 0 || (!_allowEmpty && isEmptyRecord))
                        {
                            continue;
                        }

                        _emittedRecordCount++;
                        ReportProgress(_options, _emittedRecordCount, _lineNumber - 1);
                        _visitor.Complete(fieldCount, _options.ColumnCountMismatchPolicy);
                        _currentPhysicalLineNumber = recordStartLineNumber;
                        _currentPhysicalEndLineNumber = GetCurrentPhysicalEndLineNumber(
                            recordStartLineNumber,
                            recordStartedFromPendingLine);
                        return true;
                    }
                }
                else
                {
                    lineSeparator = string.Empty;
                    readResult = CsvLineReadResult.Line;
                }

                string? line = _pendingLines.Count > 0
                    ? ReadLineWithSeparator(_lineReader, _pendingLines, out lineSeparator)
                    : fastLine;
                if (line is null)
                {
                    return false;
                }

                bool startsWithCommentCharacter = IsRawCommentLine(line, _options);
                if (TrySkipCommentRecordBeforeParsing(
                        _lineReader,
                        _pendingLines,
                        startsWithCommentCharacter,
                        line,
                        lineSeparator,
                        _options,
                        _emittedRecordCount,
                        ref _lineNumber))
                {
                    _lineNumber++;
                    continue;
                }

                if (line.IndexOf('"') < 0 && TrySplitUnquotedRecord(line, _delimiter, _trim, out string[] fields))
                {
                    _lineNumber++;
                    if (ShouldSkipCommentRecord(startsWithCommentCharacter, line, _options, _emittedRecordCount)
                        || !ShouldEmitRecord(fields, _allowEmpty))
                    {
                        continue;
                    }

                    VisitParsedFields(fields, _emittedRecordCount, null, ref _visitor);
                    _emittedRecordCount++;
                    ReportProgress(_options, _emittedRecordCount, _lineNumber - 1);
                    _visitor.Complete(fields.Length, _options.ColumnCountMismatchPolicy);
                    _currentPhysicalLineNumber = recordStartLineNumber;
                    _currentPhysicalEndLineNumber = GetCurrentPhysicalEndLineNumber(
                        recordStartLineNumber,
                        recordStartedFromPendingLine);
                    return true;
                }

                try
                {
                    if (!TryParseQuotedRecordContinuations(
                            _lineReader,
                            _pendingLines,
                            line,
                            lineSeparator,
                            _delimiter,
                            _trim,
                            _strictQuotes,
                            _quotedFields,
                            ref _lineNumber)
                        && !TryParseQuotedRecord(
                            line,
                            _delimiter,
                            _trim,
                            _strictQuotes,
                            _lineNumber,
                            _quotedFields))
                    {
                        throw new CsvParseException("Unterminated quoted field.", _lineNumber);
                    }
                }
                catch (CsvParseException exception) when (HandleParseError(_options, exception, _lineNumber))
                {
                    _lineNumber++;
                    continue;
                }

                _lineNumber++;
                if (ShouldSkipCommentRecord(startsWithCommentCharacter, line, _options, _emittedRecordCount)
                    || !ShouldEmitRecord(_quotedFields, _allowEmpty))
                {
                    continue;
                }

                VisitParsedFields(_quotedFields, _emittedRecordCount, null, ref _visitor);
                _emittedRecordCount++;
                ReportProgress(_options, _emittedRecordCount, _lineNumber - 1);
                _visitor.Complete(_quotedFields.Count, _options.ColumnCountMismatchPolicy);
                _currentPhysicalLineNumber = recordStartLineNumber;
                _currentPhysicalEndLineNumber = GetCurrentPhysicalEndLineNumber(
                    recordStartLineNumber,
                    recordStartedFromPendingLine);
                return true;
            }
        }

        private int GetCurrentPhysicalEndLineNumber(
            int recordStartLineNumber,
            bool recordStartedFromPendingLine) => Math.Max(
                recordStartLineNumber,
                recordStartedFromPendingLine
                    ? _lineNumber - 1
                    : Math.Max(
                        _lineNumber - 1,
                        _physicalLineOffset + _lineReader.PhysicalLineSeparatorsConsumed));

        public ReadOnlySpan<char> GetSpan(int ordinal) => _visitor.GetSpan(ordinal);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public string GetString(int ordinal) => _visitor.GetString(ordinal);

        public bool IsMissing(int ordinal) => _visitor.IsMissing(ordinal);

        public bool IsNull(int ordinal, string? nullValue) =>
            nullValue is not null
            && !_visitor.IsMissing(ordinal)
            && GetSpan(ordinal).SequenceEqual(nullValue.AsSpan());

        public int CopyStringValues(object[] values, int count, string? nullValue)
        {
            int valueCount = Math.Min(count, _visitor.SourceColumnCount);
            for (int index = 0; index < valueCount; index++)
            {
                values[index] = nullValue is not null && IsNull(index, nullValue)
                    ? DBNull.Value
                    : GetString(index);
            }

            for (int index = valueCount; index < count; index++)
            {
                values[index] = DBNull.Value;
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
            _lineReader.Dispose();
            _reader.Dispose();
        }
    }

    internal struct CsvDataReaderStreamRowVisitor : ICsvFieldSpanVisitor
    {
        private char[] _buffer;
        private static readonly string[] SingleCharacterStrings = CreateSingleCharacterStrings();
        private int[] _starts;
        private int[] _lengths;
        private string?[] _materialized;
        private bool _nextVisitIsUnescapedScratch;

        internal CsvDataReaderStreamRowVisitor(char[] buffer)
        {
            _buffer = buffer;
            _starts = new int[32];
            _lengths = new int[32];
            _materialized = new string?[32];
            _nextVisitIsUnescapedScratch = false;
            FieldCount = 0;
            SourceColumnCount = 0;
        }

        internal int FieldCount { get; private set; }

        internal int SourceColumnCount { get; private set; }

        internal void SetBuffer(char[] buffer)
        {
            _buffer = buffer ?? throw new ArgumentNullException(nameof(buffer));
        }

        internal void SetSourceColumnCount(int sourceColumnCount)
        {
            if (sourceColumnCount < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(sourceColumnCount));
            }

            EnsureCapacity(sourceColumnCount);
            SourceColumnCount = sourceColumnCount;
        }

        internal void Reset()
        {
            _nextVisitIsUnescapedScratch = false;
            FieldCount = 0;
        }

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            if ((uint)fieldIndex >= (uint)_starts.Length)
            {
                EnsureCapacity(fieldIndex + 1);
            }
            FieldCount = fieldIndex + 1;
            _lengths[fieldIndex] = value.Length;
            if (_nextVisitIsUnescapedScratch)
            {
                _starts[fieldIndex] = -1;
                _materialized[fieldIndex] = value.Length == 0 ? string.Empty : value.ToString();
                _nextVisitIsUnescapedScratch = false;
                return;
            }

            ref char bufferStart = ref MemoryMarshal.GetArrayDataReference(_buffer);
            ref char valueStart = ref MemoryMarshal.GetReference(value);
            nint byteOffset = Unsafe.ByteOffset(ref bufferStart, ref valueStart);
            _starts[fieldIndex] = checked((int)(byteOffset / sizeof(char)));
            _materialized[fieldIndex] = null;
        }

        public void VisitFieldRange(
            int recordIndex,
            int fieldIndex,
            char[] buffer,
            int start,
            int length)
        {
            if ((uint)fieldIndex >= (uint)_starts.Length)
            {
                EnsureCapacity(fieldIndex + 1);
            }
            FieldCount = fieldIndex + 1;
            _starts[fieldIndex] = start;
            _lengths[fieldIndex] = length;
            _materialized[fieldIndex] = null;
            _nextVisitIsUnescapedScratch = false;
        }

        internal void VisitFieldRanges(
            char[] buffer,
            int start,
            int end,
            ReadOnlySpan<int> delimiterIndexes)
        {
            int fieldCount = delimiterIndexes.Length + 1;
            EnsureCapacity(fieldCount);
            int fieldStart = start;
            for (int fieldIndex = 0; fieldIndex < delimiterIndexes.Length; fieldIndex++)
            {
                int delimiterIndex = delimiterIndexes[fieldIndex];
                _starts[fieldIndex] = fieldStart;
                _lengths[fieldIndex] = delimiterIndex - fieldStart;
                _materialized[fieldIndex] = null;
                fieldStart = delimiterIndex + 1;
            }

            int finalFieldIndex = fieldCount - 1;
            _starts[finalFieldIndex] = fieldStart;
            _lengths[finalFieldIndex] = end - fieldStart;
            _materialized[finalFieldIndex] = null;
            _nextVisitIsUnescapedScratch = false;
            FieldCount = fieldCount;
        }

        public bool TryVisitEscapedField(
            int recordIndex,
            int fieldIndex,
            ReadOnlySpan<char> escapedValue,
            int unescapedLength)
        {
            _nextVisitIsUnescapedScratch = true;
            return false;
        }

        public void VisitFieldValue(int recordIndex, int fieldIndex, string value)
        {
            if ((uint)fieldIndex >= (uint)_starts.Length)
            {
                EnsureCapacity(fieldIndex + 1);
            }
            FieldCount = fieldIndex + 1;
            _starts[fieldIndex] = -1;
            _lengths[fieldIndex] = value.Length;
            _materialized[fieldIndex] = value;
            _nextVisitIsUnescapedScratch = false;
        }

        internal void Complete(int fieldCount, CsvColumnCountMismatchPolicy mismatchPolicy)
        {
            EnsureCapacity(fieldCount);
            FieldCount = fieldCount;
            if (SourceColumnCount > 0
                && mismatchPolicy == CsvColumnCountMismatchPolicy.Strict
                && fieldCount != SourceColumnCount)
            {
                throw new CsvException(
                    $"Row contains {fieldCount} values but header defines {SourceColumnCount} columns.");
            }

            int expectedCount = SourceColumnCount > 0 ? SourceColumnCount : fieldCount;
            for (int index = fieldCount; index < expectedCount; index++)
            {
                _starts[index] = -1;
                _lengths[index] = -1;
                _materialized[index] = null;
            }
        }

        internal ReadOnlySpan<char> GetSpan(int ordinal)
        {
            ValidateOrdinal(ordinal);
            if (_lengths[ordinal] <= 0)
            {
                return ReadOnlySpan<char>.Empty;
            }

            string? materialized = _materialized[ordinal];
            return materialized is null
                ? _buffer.AsSpan(_starts[ordinal], _lengths[ordinal])
                : materialized.AsSpan();
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal string GetString(int ordinal)
        {
            ValidateOrdinal(ordinal);
            if (_lengths[ordinal] <= 0)
            {
                return string.Empty;
            }

            string? materialized = _materialized[ordinal];
            if (materialized is null)
            {
                int start = _starts[ordinal];
                int length = _lengths[ordinal];
                if (length == 1)
                {
                    char value = _buffer[start];
                    materialized = value < SingleCharacterStrings.Length
                        ? SingleCharacterStrings[value]
                        : new string(_buffer, start, length);
                }
                else
                {
                    materialized = new string(_buffer, start, length);
                }
                _materialized[ordinal] = materialized;
            }

            return materialized;
        }

        private static string[] CreateSingleCharacterStrings()
        {
            var values = new string[128];
            for (int index = 0; index < values.Length; index++)
            {
                values[index] = new string((char)index, 1);
            }
            return values;
        }

        internal bool IsMissing(int ordinal)
        {
            ValidateOrdinal(ordinal);
            return _lengths[ordinal] < 0;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void ValidateOrdinal(int ordinal)
        {
            int maximum = SourceColumnCount > 0 ? SourceColumnCount : FieldCount;
            if ((uint)ordinal >= (uint)maximum)
            {
                throw new IndexOutOfRangeException();
            }
        }

        private void EnsureCapacity(int count)
        {
            if (count <= _starts.Length)
            {
                return;
            }

            int capacity = _starts.Length;
            while (capacity < count)
            {
                capacity = checked(capacity * 2);
            }

            Array.Resize(ref _starts, capacity);
            Array.Resize(ref _lengths, capacity);
            Array.Resize(ref _materialized, capacity);
        }
    }
}
#endif
