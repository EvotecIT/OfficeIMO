#nullable enable

#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal struct CsvUtf8DataReaderRowVisitor
    {
        private readonly byte[] _buffer;
        private readonly Encoding _encoding;
        private static readonly string[] SingleCharacterStrings = CreateSingleCharacterStrings();
        private int[] _starts;
        private int[] _lengths;
        private string?[] _materialized;
        private bool _allFieldsAscii;

        internal CsvUtf8DataReaderRowVisitor(byte[] buffer, Encoding encoding)
        {
            _buffer = buffer;
            _encoding = encoding;
            _starts = new int[32];
            _lengths = new int[32];
            _materialized = new string?[32];
            FieldCount = 0;
            SourceColumnCount = 0;
        }

        internal int FieldCount { get; private set; }

        internal int SourceColumnCount { get; private set; }

        internal void SetSourceColumnCount(int sourceColumnCount)
        {
            EnsureCapacity(sourceColumnCount);
            SourceColumnCount = sourceColumnCount;
        }

        internal void Reset() => FieldCount = 0;

        internal void VisitFieldRange(int fieldIndex, int start, int length)
        {
            if ((uint)fieldIndex >= (uint)_starts.Length)
            {
                EnsureCapacity(fieldIndex + 1);
            }

            FieldCount = fieldIndex + 1;
            _starts[fieldIndex] = start;
            _lengths[fieldIndex] = length;
            _materialized[fieldIndex] = null;
        }

        internal void ShiftStarts(int offset)
        {
            for (int index = 0; index < FieldCount; index++)
            {
                _starts[index] += offset;
            }
        }

        internal void Complete(
            int fieldCount,
            CsvColumnCountMismatchPolicy mismatchPolicy,
            bool allFieldsAscii)
        {
            EnsureCapacity(fieldCount);
            FieldCount = fieldCount;
            if (SourceColumnCount > 0 &&
                mismatchPolicy == CsvColumnCountMismatchPolicy.Strict &&
                fieldCount != SourceColumnCount)
            {
                throw new CsvException(
                    $"Row contains {fieldCount} values but header defines {SourceColumnCount} columns.");
            }

            _allFieldsAscii = allFieldsAscii;

            int expectedCount = SourceColumnCount > 0 ? SourceColumnCount : fieldCount;
            for (int index = fieldCount; index < expectedCount; index++)
            {
                _starts[index] = -1;
                _lengths[index] = -1;
                _materialized[index] = null;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal string GetString(int ordinal)
        {
            ValidateOrdinal(ordinal);
            int length = _lengths[ordinal];
            if (length <= 0)
            {
                return string.Empty;
            }

            string? materialized = _materialized[ordinal];
            if (materialized is null)
            {
                int start = _starts[ordinal];
                if (length == 1 && _buffer[start] < SingleCharacterStrings.Length)
                {
                    materialized = SingleCharacterStrings[_buffer[start]];
                }
                else
                {
                    ReadOnlySpan<byte> bytes = _buffer.AsSpan(start, length);
                    materialized = _allFieldsAscii || Ascii.IsValid(bytes)
                        ? Encoding.Latin1.GetString(bytes)
                        : _encoding.GetString(bytes);
                }

                _materialized[ordinal] = materialized;
            }

            return materialized;
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

        private static string[] CreateSingleCharacterStrings()
        {
            var values = new string[128];
            for (int index = 0; index < values.Length; index++)
            {
                values[index] = new string((char)index, 1);
            }

            return values;
        }
    }
}
#endif
