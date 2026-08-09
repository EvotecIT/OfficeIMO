#nullable enable

using System.Buffers;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
#if NET8_0_OR_GREATER
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
#endif

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    private const int DefaultTextDataReaderBatchRowCapacity = 128;
    private const int MaximumTextDataReaderBatchRowCapacity = 4096;
    private const int LargeTextParallelBatchRowCapacity = 3584;
    private const int LargeTextParallelBatchThreshold = 8 * 1024 * 1024;

    internal static int GetPreferredTextParallelBatchSize(int textLength) =>
        textLength < LargeTextParallelBatchThreshold ? 2048 : LargeTextParallelBatchRowCapacity;

    private static bool CanUseTextDataReaderBatchAvx2(CsvLoadOptions options, int sourceColumnCount) =>
        Avx2.IsSupported &&
        sourceColumnCount is > 0 and <= TextQuoteAwareFieldSpanCapacity &&
        CanUseAvx2PackedDelimiter(GetDelimiterChar(options)) &&
        options.ProgressCallback is null;

    private static bool TryFillTextDataReaderBatchAvx2(
        ReadOnlySpan<char> text,
        CsvLoadOptions options,
        ref CsvTextFieldSpanReadState state,
        CsvTextDataReaderBatch batch,
        CancellationToken cancellationToken = default)
    {
        if (state.RecordsToSkip != 0 ||
            !state.UseAvx2UnquotedFastPath ||
            state.Position >= text.Length)
        {
            return false;
        }

        ThrowIfCancellationRequested(options);
        cancellationToken.ThrowIfCancellationRequested();
        batch.Reset();

        var initialPosition = state.Position;
        var position = initialPosition;
        var rowStart = position;
        var fieldStart = position;
        var fieldCount = 0;
        var firstFieldLength = 0;
        var quoteCount = 0;
        var expectedEscapedQuoteIndex = -1;
        var maximumFieldCount = 0;
        var delimiter = GetDelimiterChar(options);
        var trim = options.TrimWhitespace;
        var allowEmptyLines = options.AllowEmptyLines;
        var delimiterVector = Vector256.Create((byte)delimiter);
        var end = text.Length - 32;
        ref ushort textStart = ref Unsafe.As<char, ushort>(ref MemoryMarshal.GetReference(text));

        while (position <= end)
        {
            Vector256<byte> packedBytes;
            if (Avx512BW.IsSupported)
            {
                var values = Vector512.LoadUnsafe(ref textStart, (nuint)position);
                packedBytes = Avx512BW.ConvertToVector256ByteWithSaturation(values);
            }
            else
            {
                var values = MemoryMarshal.Cast<char, short>(text.Slice(position, 32));
                var first = Vector256.LoadUnsafe(ref MemoryMarshal.GetReference(values));
                var second = Vector256.LoadUnsafe(ref MemoryMarshal.GetReference(values.Slice(16)));
                var packed = Avx2.PackUnsignedSaturate(first, second);
                packedBytes = Vector256.AsByte(
                    Avx2.Permute4x64(Vector256.AsInt64(packed), 0b11_01_10_00));
            }

            var delimiterMask = (uint)Avx2.MoveMask(Avx2.CompareEqual(packedBytes, delimiterVector));
            var quoteMask = (uint)Avx2.MoveMask(Avx2.CompareEqual(packedBytes, QuoteByteVector));
            var carriageReturnMask = (uint)Avx2.MoveMask(
                Avx2.CompareEqual(packedBytes, CarriageReturnByteVector));
            var lineFeedMask = (uint)Avx2.MoveMask(Avx2.CompareEqual(packedBytes, LineFeedByteVector));
            var specialMask = delimiterMask | quoteMask | carriageReturnMask | lineFeedMask;

            while (specialMask != 0)
            {
                var offset = BitOperations.TrailingZeroCount(specialMask);
                var bit = 1u << offset;
                specialMask &= specialMask - 1;
                var absoluteIndex = position + offset;
                if (absoluteIndex < fieldStart)
                {
                    continue;
                }

                if ((bit & quoteMask) != 0)
                {
                    if (expectedEscapedQuoteIndex >= 0)
                    {
                        if (absoluteIndex != expectedEscapedQuoteIndex)
                        {
                            return CompleteOrRejectTextDataReaderBatch(
                                ref state,
                                batch,
                                initialPosition,
                                rowStart,
                                maximumFieldCount);
                        }

                        expectedEscapedQuoteIndex = -1;
                    }
                    else if ((quoteCount & 1) != 0 &&
                        absoluteIndex + 1 < text.Length &&
                        text[absoluteIndex + 1] == '"')
                    {
                        if (offset != 31)
                        {
                            specialMask &= ~(bit << 1);
                            quoteCount += 2;
                            continue;
                        }

                        expectedEscapedQuoteIndex = absoluteIndex + 1;
                    }
                    else if (quoteCount != 0 && (quoteCount & 1) == 0)
                    {
                        return CompleteOrRejectTextDataReaderBatch(
                            ref state,
                            batch,
                            initialPosition,
                            rowStart,
                            maximumFieldCount);
                    }

                    quoteCount++;
                    continue;
                }

                if ((quoteCount & 1) != 0)
                {
                    continue;
                }

                if ((bit & delimiterMask) != 0)
                {
                    if (!TryAddTextDataReaderBatchField(
                        text,
                        trim,
                        batch,
                        fieldStart,
                        absoluteIndex,
                        quoteCount,
                        fieldCount,
                        out var fieldLength))
                    {
                        return CompleteOrRejectTextDataReaderBatch(
                            ref state,
                            batch,
                            initialPosition,
                            rowStart,
                            maximumFieldCount);
                    }

                    if (fieldCount == 0)
                    {
                        firstFieldLength = fieldLength;
                    }

                    fieldCount++;
                    fieldStart = absoluteIndex + 1;
                    quoteCount = 0;
                    expectedEscapedQuoteIndex = -1;
                    continue;
                }

                if ((bit & (carriageReturnMask | lineFeedMask)) == 0)
                {
                    continue;
                }

                if (!TryAddTextDataReaderBatchField(
                    text,
                    trim,
                    batch,
                    fieldStart,
                    absoluteIndex,
                    quoteCount,
                    fieldCount,
                    out var finalFieldLength))
                {
                    return CompleteOrRejectTextDataReaderBatch(
                        ref state,
                        batch,
                        initialPosition,
                        rowStart,
                        maximumFieldCount);
                }

                if (fieldCount == 0)
                {
                    firstFieldLength = finalFieldLength;
                }

                fieldCount++;
                var nextPosition = absoluteIndex + 1;
                if ((bit & carriageReturnMask) != 0 &&
                    nextPosition < text.Length &&
                    text[nextPosition] == '\n')
                {
                    nextPosition++;
                }

                maximumFieldCount = Math.Max(maximumFieldCount, fieldCount);
                var isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
                if (allowEmptyLines || !isEmptyRecord)
                {
                    batch.CompleteRow(fieldCount, options.ColumnCountMismatchPolicy);
                    if (batch.IsFull)
                    {
                        CommitTextDataReaderBatch(ref state, batch, nextPosition, maximumFieldCount);
                        return true;
                    }
                }
                else
                {
                    batch.DiscardPendingRow();
                }

                rowStart = nextPosition;
                fieldStart = nextPosition;
                fieldCount = 0;
                firstFieldLength = 0;
                quoteCount = 0;
                expectedEscapedQuoteIndex = -1;
            }

            position += 32;
            if (((position - initialPosition) & 4095) == 0)
            {
                ThrowIfCancellationRequested(options);
                cancellationToken.ThrowIfCancellationRequested();
            }
        }

        while (position < text.Length)
        {
            var value = text[position];
            if (value == '"')
            {
                if (expectedEscapedQuoteIndex >= 0)
                {
                    if (position != expectedEscapedQuoteIndex)
                    {
                        return CompleteOrRejectTextDataReaderBatch(
                            ref state,
                            batch,
                            initialPosition,
                            rowStart,
                            maximumFieldCount);
                    }

                    expectedEscapedQuoteIndex = -1;
                }
                else if ((quoteCount & 1) != 0 &&
                    position + 1 < text.Length &&
                    text[position + 1] == '"')
                {
                    quoteCount += 2;
                    position += 2;
                    continue;
                }
                else if (quoteCount != 0 && (quoteCount & 1) == 0)
                {
                    return CompleteOrRejectTextDataReaderBatch(
                        ref state,
                        batch,
                        initialPosition,
                        rowStart,
                        maximumFieldCount);
                }

                quoteCount++;
                position++;
                continue;
            }

            if ((quoteCount & 1) == 0 && value == delimiter)
            {
                if (!TryAddTextDataReaderBatchField(
                    text,
                    trim,
                    batch,
                    fieldStart,
                    position,
                    quoteCount,
                    fieldCount,
                    out var fieldLength))
                {
                    return CompleteOrRejectTextDataReaderBatch(
                        ref state,
                        batch,
                        initialPosition,
                        rowStart,
                        maximumFieldCount);
                }

                if (fieldCount == 0)
                {
                    firstFieldLength = fieldLength;
                }

                fieldCount++;
                fieldStart = ++position;
                quoteCount = 0;
                expectedEscapedQuoteIndex = -1;
                continue;
            }

            if ((quoteCount & 1) == 0 && (value == '\r' || value == '\n'))
            {
                if (!TryAddTextDataReaderBatchField(
                    text,
                    trim,
                    batch,
                    fieldStart,
                    position,
                    quoteCount,
                    fieldCount,
                    out var finalFieldLength))
                {
                    return CompleteOrRejectTextDataReaderBatch(
                        ref state,
                        batch,
                        initialPosition,
                        rowStart,
                        maximumFieldCount);
                }

                if (fieldCount == 0)
                {
                    firstFieldLength = finalFieldLength;
                }

                fieldCount++;
                var nextPosition = position + 1;
                if (value == '\r' && nextPosition < text.Length && text[nextPosition] == '\n')
                {
                    nextPosition++;
                }

                maximumFieldCount = Math.Max(maximumFieldCount, fieldCount);
                var isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
                if (allowEmptyLines || !isEmptyRecord)
                {
                    batch.CompleteRow(fieldCount, options.ColumnCountMismatchPolicy);
                    if (batch.IsFull)
                    {
                        CommitTextDataReaderBatch(ref state, batch, nextPosition, maximumFieldCount);
                        return true;
                    }
                }
                else
                {
                    batch.DiscardPendingRow();
                }

                position = nextPosition;
                rowStart = position;
                fieldStart = position;
                fieldCount = 0;
                firstFieldLength = 0;
                quoteCount = 0;
                expectedEscapedQuoteIndex = -1;
                continue;
            }

            position++;
        }

        if (fieldStart < text.Length || fieldCount != 0)
        {
            if (!TryAddTextDataReaderBatchField(
                text,
                trim,
                batch,
                fieldStart,
                text.Length,
                quoteCount,
                fieldCount,
                out var finalFieldLength))
            {
                return CompleteOrRejectTextDataReaderBatch(
                    ref state,
                    batch,
                    initialPosition,
                    rowStart,
                    maximumFieldCount);
            }

            if (fieldCount == 0)
            {
                firstFieldLength = finalFieldLength;
            }

            fieldCount++;
            maximumFieldCount = Math.Max(maximumFieldCount, fieldCount);
            var isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
            if (allowEmptyLines || !isEmptyRecord)
            {
                batch.CompleteRow(fieldCount, options.ColumnCountMismatchPolicy);
            }
            else
            {
                batch.DiscardPendingRow();
            }
        }

        CommitTextDataReaderBatch(ref state, batch, text.Length, maximumFieldCount);
        return batch.RowCount != 0;
    }

    private static bool TryAddTextDataReaderBatchField(
        ReadOnlySpan<char> text,
        bool trim,
        CsvTextDataReaderBatch batch,
        int start,
        int end,
        int quoteCount,
        int fieldIndex,
        out int fieldLength)
    {
        if (trim)
        {
            while (start < end && IsTextDataReaderWhitespace(text[start]))
            {
                start++;
            }
            while (end > start && IsTextDataReaderWhitespace(text[end - 1]))
            {
                end--;
            }
        }

        if (quoteCount == 0)
        {
            fieldLength = end - start;
            return batch.TrySetPendingField(fieldIndex, start, fieldLength, escapedSourceLength: 0);
        }

        var rawLength = end - start;
        if ((quoteCount & 1) != 0 ||
            rawLength < 2 ||
            text[start] != '"' ||
            text[end - 1] != '"')
        {
            fieldLength = 0;
            return false;
        }

        var valueStart = start + 1;
        var valueLength = rawLength - 2;
        var escapedQuoteCount = quoteCount - 2;
        if ((escapedQuoteCount & 1) != 0)
        {
            fieldLength = 0;
            return false;
        }

        fieldLength = valueLength - (escapedQuoteCount / 2);
        return batch.TrySetPendingField(
            fieldIndex,
            valueStart,
            fieldLength,
            escapedQuoteCount == 0 ? 0 : valueLength);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool IsTextDataReaderWhitespace(char value)
    {
        if (value == ' ')
        {
            return true;
        }

        if (value <= '\r')
        {
            return value is '\t' or '\n' or '\v' or '\f' or '\r';
        }

        return value > '\u007f' && char.IsWhiteSpace(value);
    }

    private static bool CompleteOrRejectTextDataReaderBatch(
        ref CsvTextFieldSpanReadState state,
        CsvTextDataReaderBatch batch,
        int initialPosition,
        int rejectedRowStart,
        int maximumFieldCount)
    {
        batch.DiscardPendingRow();
        if (batch.RowCount == 0)
        {
            batch.Reset();
            state.Position = initialPosition;
            return false;
        }

        CommitTextDataReaderBatch(ref state, batch, rejectedRowStart, maximumFieldCount);
        return true;
    }

    private static void CommitTextDataReaderBatch(
        ref CsvTextFieldSpanReadState state,
        CsvTextDataReaderBatch batch,
        int position,
        int maximumFieldCount)
    {
        state.Position = position;
        state.RecordIndex += batch.RowCount;
        state.EmittedRecordCount += batch.RowCount;
        state.LineNumber += batch.RowCount;
        var requiredFieldCapacity = GetTextDelimiterIndexCapacity(maximumFieldCount);
        if (requiredFieldCapacity > state.UnquotedDelimiterIndexCapacity)
        {
            state.UnquotedDelimiterIndexCapacity = requiredFieldCapacity;
        }
    }

    internal sealed class CsvTextDataReaderBatch : ICsvDataReaderTextRowSource, ICsvDataReaderParallelBatchInfo
    {
        private readonly string _text;
        private readonly int _sourceColumnCount;
        private readonly CultureInfo _culture;
        private readonly IReadOnlyList<string>? _dateTimeFormats;
        private readonly int[] _starts;
        // The decoded and escaped-source lengths are independently bounded by ushort.MaxValue.
        // Packing them keeps the hot metadata footprint to two 32-bit arrays; oversized fields
        // reject this optional batch path and resume through the general parser without truncation.
        private readonly uint[] _lengths;
        private string?[]? _materialized;
        private string? _nullValue;
        private readonly int[] _fieldCounts;
        private readonly int _rowCapacity;
        private readonly int _fieldCapacity;
        private int _currentRow = -1;
        private int _currentFieldOffset;
        private bool _strictColumnCount;
        private bool _disposed;

        internal CsvTextDataReaderBatch(
            string text,
            int sourceColumnCount,
            CultureInfo culture,
            IReadOnlyList<string>? dateTimeFormats,
            int preferredRowCapacity = DefaultTextDataReaderBatchRowCapacity)
        {
            _text = text;
            _sourceColumnCount = sourceColumnCount;
            _culture = culture;
            _dateTimeFormats = dateTimeFormats;
            _rowCapacity = Math.Max(
                1,
                Math.Min(MaximumTextDataReaderBatchRowCapacity, preferredRowCapacity));
            _fieldCapacity = checked(_rowCapacity * sourceColumnCount);
            _starts = ArrayPool<int>.Shared.Rent(_fieldCapacity);
            _lengths = ArrayPool<uint>.Shared.Rent(_fieldCapacity);
            _fieldCounts = ArrayPool<int>.Shared.Rent(_rowCapacity);
        }

        internal int SourceColumnCount => _sourceColumnCount;

        internal CultureInfo Culture => _culture;

        internal IReadOnlyList<string>? DateTimeFormats => _dateTimeFormats;

        internal int RowCapacity => _rowCapacity;

        internal bool HasStarted => _currentRow >= 0;

        internal int RowCount { get; private set; }

        int ICsvDataReaderParallelBatchInfo.RowCount => RowCount;

        internal int CurrentFieldCount => _fieldCounts[_currentRow];

        internal bool IsFull => RowCount == _rowCapacity;

        internal void Reset()
        {
            RowCount = 0;
            _currentRow = -1;
            _currentFieldOffset = 0;
            _strictColumnCount = false;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal bool MoveNext()
        {
            if (_currentRow + 1 >= RowCount)
            {
                return false;
            }

            _currentRow++;
            _currentFieldOffset = _currentRow * _sourceColumnCount;
            if (_strictColumnCount && CurrentFieldCount != _sourceColumnCount)
            {
                throw new CsvException(
                    $"Row contains {CurrentFieldCount} values but header defines {_sourceColumnCount} columns.");
            }
            return true;
        }

        internal bool TrySetPendingField(
            int fieldIndex,
            int start,
            int length,
            int escapedSourceLength)
        {
            if ((uint)fieldIndex >= (uint)_sourceColumnCount)
            {
                return true;
            }
            if ((uint)length > ushort.MaxValue || (uint)escapedSourceLength > ushort.MaxValue)
            {
                return false;
            }

            var index = RowCount * _sourceColumnCount + fieldIndex;
            _starts[index] = start;
            _lengths[index] = (ushort)length | ((uint)(ushort)escapedSourceLength << 16);
            if (_materialized is not null)
            {
                _materialized[index] = null;
            }
            return true;
        }

        internal void CompleteRow(
            int fieldCount,
            CsvColumnCountMismatchPolicy mismatchPolicy)
        {
            _strictColumnCount |= mismatchPolicy == CsvColumnCountMismatchPolicy.Strict;
            _fieldCounts[RowCount] = fieldCount;
            RowCount++;
        }

        internal void DiscardPendingRow()
        {
            var start = RowCount * _sourceColumnCount;
            if (_materialized is not null)
            {
                Array.Clear(_materialized, start, _sourceColumnCount);
            }
        }

        internal bool IsMissing(int ordinal)
        {
            ValidateOrdinal(ordinal);
            return ordinal >= CurrentFieldCount;
        }

        internal void SetNullValue(string? nullValue) => _nullValue = nullValue;

        internal bool IsConfiguredNull(int ordinal) => IsNull(ordinal, _nullValue);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool Read() => MoveNext();

        public bool IsNull(int ordinal, string? nullValue) =>
            !IsMissing(ordinal) &&
            nullValue is not null &&
            GetSpan(ordinal).SequenceEqual(nullValue.AsSpan());

        public int CopyStringValues(object[] values, int count, string? nullValue)
        {
            int valueCount = Math.Min(count, _sourceColumnCount);
            for (int index = 0; index < valueCount; index++)
            {
                values[index] = IsNull(index, nullValue) ? DBNull.Value : GetString(index);
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
            ArrayPool<int>.Shared.Return(_starts);
            ArrayPool<uint>.Shared.Return(_lengths);
            if (_materialized is not null)
            {
                ArrayPool<string?>.Shared.Return(_materialized, clearArray: true);
                _materialized = null;
            }
            ArrayPool<int>.Shared.Return(_fieldCounts);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public ReadOnlySpan<char> GetSpan(int ordinal)
        {
            ValidateOrdinal(ordinal);
            if (ordinal >= CurrentFieldCount)
            {
                return ReadOnlySpan<char>.Empty;
            }

            var index = GetCurrentFieldIndex(ordinal);
            uint lengths = _lengths[index];
            if (GetEscapedSourceLength(lengths) != 0)
            {
                return GetString(ordinal).AsSpan();
            }

            return _text.AsSpan(_starts[index], GetLength(lengths));
        }

        public string GetString(int ordinal)
        {
            ValidateOrdinal(ordinal);
            if (ordinal >= CurrentFieldCount)
            {
                return string.Empty;
            }

            var index = GetCurrentFieldIndex(ordinal);
            var length = GetLength(_lengths[index]);
            if (length == 0)
            {
                return string.Empty;
            }

            string?[] materializedValues = GetOrCreateMaterializedValues();
            var materialized = materializedValues[index];
            if (materialized is not null)
            {
                return materialized;
            }

            materialized = MaterializeString(index, length);
            materializedValues[index] = materialized;
            return materialized;
        }

        bool ICsvDataReaderTextRowSource.IsMissing(int ordinal) => IsMissing(ordinal);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal string MaterializeString(int ordinal)
        {
            ValidateOrdinal(ordinal);
            if (ordinal >= CurrentFieldCount)
            {
                return string.Empty;
            }

            int index = GetCurrentFieldIndex(ordinal);
            int length = GetLength(_lengths[index]);
            return length == 0 ? string.Empty : MaterializeString(index, length);
        }

        private string MaterializeString(int index, int length)
        {
            uint lengths = _lengths[index];
            var escapedSourceLength = GetEscapedSourceLength(lengths);
            if (escapedSourceLength == 0)
            {
                return _text.Substring(_starts[index], length);
            }

            return string.Create(
                length,
                (Text: _text, Start: _starts[index], Length: escapedSourceLength),
                static (destination, sourceState) =>
                {
                    var source = sourceState.Text.AsSpan(sourceState.Start, sourceState.Length);
                    var sourceIndex = 0;
                    var destinationIndex = 0;
                    while (sourceIndex < source.Length)
                    {
                        var value = source[sourceIndex++];
                        destination[destinationIndex++] = value;
                        if (value == '"' &&
                            sourceIndex < source.Length &&
                            source[sourceIndex] == '"')
                        {
                            sourceIndex++;
                        }
                    }
                });
        }

        private string?[] GetOrCreateMaterializedValues()
        {
            if (_materialized is not null)
            {
                return _materialized;
            }

            string?[] values = ArrayPool<string?>.Shared.Rent(_fieldCapacity);
            // ArrayPool does not promise cleared references. Stale strings would be observable as
            // field values, so clear the portion this batch can address before publishing it.
            Array.Clear(values, 0, _fieldCapacity);
            _materialized = values;
            return values;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static int GetLength(uint lengths) => (ushort)lengths;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static int GetEscapedSourceLength(uint lengths) => (ushort)(lengths >> 16);

        private int GetCurrentFieldIndex(int ordinal) => _currentFieldOffset + ordinal;

        private void ValidateOrdinal(int ordinal)
        {
            if ((uint)ordinal >= (uint)_sourceColumnCount)
            {
                throw new IndexOutOfRangeException();
            }
        }

    }
#endif
}
