#if !NET8_0_OR_GREATER
using System.Buffers;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// .NET Framework and .NET Standard 2.0 support for the UTF-8 stream reader. Newer targets use the
// platform types these stand in for, and the vectorized block scanner instead of the word scanner.
namespace System.Buffers
{
    /// <summary>The bytes the UTF-8 reader searches for, matched eight at a time.</summary>
    internal static class SearchValues
    {
        internal static SearchValues<byte> Create(byte[] values) => new(values);
    }

    /// <summary>Up to four bytes to find, each repeated across a 64-bit word.</summary>
    internal sealed class SearchValues<T>
    {
        internal SearchValues(byte[] values)
        {
            if (values.Length is 0 or > 4)
            {
                throw new ArgumentOutOfRangeException(nameof(values));
            }

            Values = values;
            First = Broadcast(values[0]);
            Second = Broadcast(values[Math.Min(1, values.Length - 1)]);
            Third = Broadcast(values[Math.Min(2, values.Length - 1)]);
            Fourth = Broadcast(values[values.Length - 1]);
        }

        internal byte[] Values { get; }
        internal ulong First { get; }
        internal ulong Second { get; }
        internal ulong Third { get; }
        internal ulong Fourth { get; }

        private static ulong Broadcast(byte value) => value * 0x0101010101010101UL;
    }
}

namespace System.Text
{
    /// <summary>The ASCII checks the UTF-8 reader uses.</summary>
    internal static class Ascii
    {
        internal static bool IsValid(ReadOnlySpan<byte> value)
        {
            ref byte start = ref MemoryMarshal.GetReference(value);
            int index = 0;
            for (; index <= value.Length - 8; index += 8)
            {
                if ((Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref start, index)) & 0x8080808080808080UL) != 0)
                {
                    return false;
                }
            }

            for (; index < value.Length; index++)
            {
                if (value[index] >= 0x80)
                {
                    return false;
                }
            }

            return true;
        }

        internal static bool Equals(ReadOnlySpan<byte> left, string right)
        {
            if (left.Length != right.Length)
            {
                return false;
            }

            for (int index = 0; index < left.Length; index++)
            {
                if (left[index] >= 0x80 || left[index] != right[index])
                {
                    return false;
                }
            }

            return true;
        }

        internal static bool EqualsIgnoreCase(ReadOnlySpan<byte> left, ReadOnlySpan<byte> right)
        {
            if (left.Length != right.Length)
            {
                return false;
            }

            for (int index = 0; index < left.Length; index++)
            {
                byte first = left[index];
                byte second = right[index];
                if (first >= 0x80 || second >= 0x80 || ToLower(first) != ToLower(second))
                {
                    return false;
                }
            }

            return true;
        }

        private static byte ToLower(byte value) => value is >= (byte)'A' and <= (byte)'Z' ? (byte)(value | 0x20) : value;
    }

    /// <summary>A Unicode scalar decoded from UTF-8 as <c>System.Text.Rune</c> decodes it.</summary>
    internal readonly struct Rune
    {
        private Rune(int value) => Value = value;

        internal int Value { get; }

        internal static bool IsWhiteSpace(Rune value) => value.Value <= char.MaxValue && char.IsWhiteSpace((char)value.Value);

        /// <summary>
        /// Decodes the first scalar. Invalid input consumes its longest valid prefix, at least one
        /// byte; a valid but incomplete sequence at the end needs more data.
        /// </summary>
        internal static OperationStatus DecodeFromUtf8(ReadOnlySpan<byte> source, out Rune result, out int bytesConsumed)
        {
            result = new Rune(0xFFFD);
            if (source.IsEmpty)
            {
                bytesConsumed = 0;
                return OperationStatus.NeedMoreData;
            }

            byte lead = source[0];
            if (lead < 0x80)
            {
                result = new Rune(lead);
                bytesConsumed = 1;
                return OperationStatus.Done;
            }

            int length;
            int value;
            if (lead is >= 0xC2 and <= 0xDF)
            {
                length = 2;
                value = lead & 0x1F;
            }
            else if (lead is >= 0xE0 and <= 0xEF)
            {
                length = 3;
                value = lead & 0x0F;
            }
            else if (lead is >= 0xF0 and <= 0xF4)
            {
                length = 4;
                value = lead & 0x07;
            }
            else
            {
                bytesConsumed = 1;
                return OperationStatus.InvalidData;
            }

            for (int index = 1; index < length; index++)
            {
                if (index == source.Length)
                {
                    bytesConsumed = index;
                    return OperationStatus.NeedMoreData;
                }

                byte next = source[index];
                byte lowest = 0x80;
                byte highest = 0xBF;
                if (index == 1)
                {
                    // Reject overlong forms, surrogates and scalars above U+10FFFF.
                    if (lead == 0xE0) lowest = 0xA0;
                    else if (lead == 0xED) highest = 0x9F;
                    else if (lead == 0xF0) lowest = 0x90;
                    else if (lead == 0xF4) highest = 0x8F;
                }

                if (next < lowest || next > highest)
                {
                    bytesConsumed = index;
                    return OperationStatus.InvalidData;
                }

                value = (value << 6) | (next & 0x3F);
            }

            result = new Rune(value);
            bytesConsumed = length;
            return OperationStatus.Done;
        }

        /// <summary>Decodes the last scalar; anything but a complete valid sequence is invalid.</summary>
        internal static OperationStatus DecodeLastFromUtf8(ReadOnlySpan<byte> source, out Rune result, out int bytesConsumed)
        {
            int start = source.Length - 1;
            while (start > 0 && start > source.Length - 4 && (source[start] & 0xC0) == 0x80)
            {
                start--;
            }

            if (start >= 0 && DecodeFromUtf8(source.Slice(start), out result, out bytesConsumed) == OperationStatus.Done &&
                bytesConsumed == source.Length - start)
            {
                return OperationStatus.Done;
            }

            result = new Rune(0xFFFD);
            bytesConsumed = 1;
            return OperationStatus.InvalidData;
        }
    }
}

namespace OfficeIMO.CSV
{
    internal static class CsvDownLevelSpanExtensions
    {
        /// <summary>Finds the first of the searched bytes eight bytes at a time.</summary>
        internal static int IndexOfAny(this Span<byte> span, SearchValues<byte> values) =>
            IndexOfAny((ReadOnlySpan<byte>)span, values);

        internal static int IndexOfAny(this ReadOnlySpan<byte> span, SearchValues<byte> values)
        {
            ref byte start = ref MemoryMarshal.GetReference(span);
            int index = 0;
            if (BitConverter.IsLittleEndian)
            {
                ulong first = values.First, second = values.Second, third = values.Third, fourth = values.Fourth;
                for (; index <= span.Length - 8; index += 8)
                {
                    ulong word = Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref start, index));
                    ulong found = ZeroBytes(word ^ first) | ZeroBytes(word ^ second) | ZeroBytes(word ^ third) | ZeroBytes(word ^ fourth);
                    if (found != 0)
                    {
                        // The lowest flag is exact; bytes below it each contribute one to the sum.
                        return index + (int)((((found & (0UL - found)) - 1UL) & 0x0101010101010101UL) * 0x0101010101010101UL >> 56) - 1;
                    }
                }
            }

            byte[] searched = values.Values;
            for (; index < span.Length; index++)
            {
                if (Array.IndexOf(searched, Unsafe.Add(ref start, index)) >= 0)
                {
                    return index;
                }
            }

            return -1;
        }

        internal static bool Contains(this Span<byte> span, byte value) => span.IndexOf(value) >= 0;

        internal static bool Contains(this ReadOnlySpan<byte> span, byte value) => span.IndexOf(value) >= 0;

        /// <summary>Flags the zero bytes of a word; only the lowest flag is guaranteed exact.</summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static ulong ZeroBytes(ulong word) => (word - 0x0101010101010101UL) & ~word & 0x8080808080808080UL;

        internal static string GetString(this Encoding encoding, ReadOnlySpan<byte> bytes)
        {
            byte[] rented = ArrayPool<byte>.Shared.Rent(Math.Max(bytes.Length, 1));
            try
            {
                bytes.CopyTo(rented);
                return encoding.GetString(rented, 0, bytes.Length);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(rented);
            }
        }
    }

    internal static partial class CsvParser
    {
        internal sealed partial class CsvUtf8StreamDataReaderRowSource
        {
            private const ulong RepeatedOnes = 0x0101010101010101UL;
            private const ulong LowSevenBits = 0x7F7F7F7F7F7F7F7FUL;

            /// <summary>
            /// Scans a record without quotes eight bytes at a time and visits its fields directly.
            /// A quote, trimming, or fewer than eight bytes left at the end of input leave the record
            /// to the exact scanner, as <see cref="RecordScanResult.Irregular"/>.
            /// </summary>
            private RecordScanResult ScanRecordWords(int recordStart, out RecordScan scan)
            {
                scan = default;
                _visitor.Reset();
                byte[] buffer = _buffer;
                int length = _length;
                int position = recordStart;
                int fieldStart = recordStart;
                int fieldIndex = 0;
                ulong high = 0;
                ulong delimiters = _delimiter * RepeatedOnes;
                // When quotes are data the quote pattern repeats the delimiter, so no quote is flagged.
                ulong quotes = (_quotesAreData ? _delimiter : (byte)'"') * RepeatedOnes;
                ref byte data = ref buffer[0];
                while (length - position >= 8)
                {
                    ulong word = Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref data, position));
                    ulong special = MatchingBytes(word ^ delimiters) | MatchingBytes(word ^ quotes) |
                        MatchingBytes(word ^ ((byte)'\r' * RepeatedOnes)) | MatchingBytes(word ^ ((byte)'\n' * RepeatedOnes));
                    while (special != 0)
                    {
                        int offset = LowestFlaggedByte(special);
                        int index = position + offset;
                        byte value = buffer[index];
                        if (value == _delimiter)
                        {
                            _visitor.VisitFieldRange(fieldIndex++, fieldStart, index - fieldStart);
                            fieldStart = index + 1;
                        }
                        else if (value == (byte)'"' && !_quotesAreData)
                        {
                            return RecordScanResult.Irregular;
                        }
                        else
                        {
                            _visitor.VisitFieldRange(fieldIndex, fieldStart, index - fieldStart);
                            high |= word & ((1UL << (offset * 8)) - 1UL);
                            scan.End = index;
                            scan.Next = index + 1;
                            scan.FieldCount = fieldIndex + 1;
                            scan.Terminator = value;
                            scan.Ascii = (high & 0x8080808080808080UL) == 0 ? AsciiState.Ascii : AsciiState.NonAscii;
                            scan.NoEscapedQuotes = true;
                            return RecordScanResult.Complete;
                        }

                        special &= special - 1UL;
                    }

                    high |= word;
                    position += 8;
                }

                return _endOfStream ? RecordScanResult.Irregular : RecordScanResult.NeedMoreData;
            }

            /// <summary>Sets the high bit of exactly the zero bytes of a word.</summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static ulong MatchingBytes(ulong word) => ~(((word & LowSevenBits) + LowSevenBits) | word | LowSevenBits);

            /// <summary>The index of the lowest byte whose high bit is set.</summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static int LowestFlaggedByte(ulong flags) =>
                (int)(((((flags & (0UL - flags)) >> 7) - 1UL) & RepeatedOnes) * RepeatedOnes >> 56);
        }

        private static CsvStringCache? RentStringCacheWorkspace(CsvParsingSettings options) => CreateStringCache(options);

        private static void ReturnStringCacheWorkspace(CsvStringCache? cache)
        {
        }

        /// <summary>
        /// Rows after the UTF-8 reader hands over at a record boundary. The canonical record parser
        /// continues from that record's line and emitted record count.
        /// </summary>
        internal sealed class CsvStreamDataReaderRowSource : ICsvDataReaderHeaderRowSource, ICsvDataReaderPositionSource
        {
            private readonly TextReader _reader;
            private readonly CsvParsingSettings _options;
            private readonly IEnumerator<IReadOnlyList<string>> _records;
            private IReadOnlyList<string> _current = Array.Empty<string>();
            private int _sourceColumnCount;
            private bool _disposed;

            internal CsvStreamDataReaderRowSource(
                TextReader reader,
                CsvParsingSettings options,
                int initialEmittedRecordCount = 0,
                int physicalLineOffset = 0,
                CsvStringCache? stringCache = null)
            {
                _reader = reader ?? throw new ArgumentNullException(nameof(reader));
                _options = options ?? throw new ArgumentNullException(nameof(options));
                _records = ParseReusable(reader, options, physicalLineOffset + 1, initialEmittedRecordCount, stringCache).GetEnumerator();
            }

            public int FieldCount => _current.Count;

            int? ICsvDataReaderPositionSource.CurrentPhysicalLineNumber => null;

            int? ICsvDataReaderPositionSource.CurrentPhysicalEndLineNumber => null;

            public void SetSourceColumnCount(int sourceColumnCount)
            {
                if (sourceColumnCount < 0)
                {
                    throw new ArgumentOutOfRangeException(nameof(sourceColumnCount));
                }

                _sourceColumnCount = sourceColumnCount;
            }

            public bool Read() => Read(_options.CancellationToken);

            public bool Read(CancellationToken cancellationToken)
            {
                if (_disposed)
                {
                    throw new ObjectDisposedException(GetType().FullName);
                }

                cancellationToken.ThrowIfCancellationRequested();
                if (!_records.MoveNext())
                {
                    _current = Array.Empty<string>();
                    return false;
                }

                _current = _records.Current;
                if (_sourceColumnCount > 0 &&
                    _options.ColumnCountMismatchPolicy == CsvColumnCountMismatchPolicy.Strict &&
                    _current.Count != _sourceColumnCount)
                {
                    throw new CsvException(
                        $"Row contains {_current.Count} values but header defines {_sourceColumnCount} columns.");
                }

                return true;
            }

            public ReadOnlySpan<char> GetSpan(int ordinal) => GetString(ordinal).AsSpan();

            public string GetString(int ordinal)
            {
                ValidateOrdinal(ordinal);
                return ordinal < _current.Count ? _current[ordinal] : string.Empty;
            }

            public bool IsMissing(int ordinal)
            {
                ValidateOrdinal(ordinal);
                return ordinal >= _current.Count;
            }

            public bool IsNull(int ordinal, string? nullValue) =>
                nullValue is not null && !IsMissing(ordinal) && string.Equals(GetString(ordinal), nullValue, StringComparison.Ordinal);

            public int CopyStringValues(object[] values, int count, string? nullValue)
            {
                int valueCount = Math.Min(count, _sourceColumnCount > 0 ? _sourceColumnCount : _current.Count);
                for (int index = 0; index < valueCount; index++)
                {
                    values[index] = nullValue is not null && IsNull(index, nullValue) ? DBNull.Value : GetString(index);
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
                try
                {
                    _records.Dispose();
                }
                finally
                {
                    _reader.Dispose();
                }
            }

            private void ValidateOrdinal(int ordinal)
            {
                int maximum = _sourceColumnCount > 0 ? _sourceColumnCount : _current.Count;
                if ((uint)ordinal >= (uint)maximum)
                {
                    throw new IndexOutOfRangeException();
                }
            }
        }
    }
}
#endif
