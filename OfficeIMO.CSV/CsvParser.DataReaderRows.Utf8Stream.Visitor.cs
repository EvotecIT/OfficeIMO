#if NET8_0_OR_GREATER
using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
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
                        ? CreateAsciiString(_buffer, start, length)
                        : _encoding.GetString(bytes);
                }

                _materialized[ordinal] = materialized;
            }

            return materialized;
        }

        // Only validated ASCII reaches this method. The first and last bounded loads
        // overlap for short tails, so every UTF-16 character is initialized without padding.
        private static string CreateAsciiString(byte[] buffer, int start, int length)
        {
            if (!BitConverter.IsLittleEndian || !Vector128.IsHardwareAccelerated || length < 2 || length >= 16)
            {
                return Encoding.Latin1.GetString(buffer.AsSpan(start, length));
            }

            return string.Create(length, (Buffer: buffer, Start: start), static (characters, state) =>
            {
                ReadOnlySpan<byte> bytes = state.Buffer.AsSpan(state.Start, characters.Length);
                Span<byte> destination = MemoryMarshal.AsBytes(characters);
                int length = bytes.Length;
                if (length >= 8)
                {
                    WidenEight(bytes, destination);
                    WidenEight(bytes.Slice(length - 8), destination.Slice((length - 8) * 2));
                }
                else if (length >= 4)
                {
                    WidenFour(bytes, destination);
                    WidenFour(bytes.Slice(length - 4), destination.Slice((length - 4) * 2));
                }
                else
                {
                    WidenTwo(bytes, destination);
                    WidenTwo(bytes.Slice(length - 2), destination.Slice((length - 2) * 2));
                }
            });
        }

        private static void WidenEight(ReadOnlySpan<byte> bytes, Span<byte> destination)
        {
            ulong word = BinaryPrimitives.ReadUInt64LittleEndian(bytes);
            Vector128.WidenLower(Vector128.CreateScalarUnsafe(word).AsByte()).AsByte().CopyTo(destination);
        }

        private static void WidenFour(ReadOnlySpan<byte> bytes, Span<byte> destination)
        {
            ulong word = BinaryPrimitives.ReadUInt32LittleEndian(bytes);
            word = (word | (word << 16)) & 0x0000ffff0000ffffUL;
            word = (word | (word << 8)) & 0x00ff00ff00ff00ffUL;
            BinaryPrimitives.WriteUInt64LittleEndian(destination, word);
        }

        private static void WidenTwo(ReadOnlySpan<byte> bytes, Span<byte> destination)
        {
            uint word = BinaryPrimitives.ReadUInt16LittleEndian(bytes);
            BinaryPrimitives.WriteUInt32LittleEndian(destination, (word | (word << 8)) & 0x00ff00ffU);
        }

        /// <summary>Completes requested per-read string sharing before exposing the row.</summary>
        internal void InternFields(Dictionary<string, string>? cache)
        {
            if (cache is null) return;
            int count = SourceColumnCount > 0 ? Math.Min(FieldCount, SourceColumnCount) : FieldCount;
#if NET9_0_OR_GREATER
            Span<char> characters = stackalloc char[128];
#endif
            for (int ordinal = 0; ordinal < count; ordinal++)
            {
#if NET9_0_OR_GREATER
                int length = _lengths[ordinal];
                if (length > 0 && length <= characters.Length)
                {
                    ReadOnlySpan<byte> bytes = _buffer.AsSpan(_starts[ordinal], length);
                    if (_allFieldsAscii || Ascii.IsValid(bytes))
                    {
                        int written = Encoding.Latin1.GetChars(bytes, characters);
                        _materialized[ordinal] = ReuseFieldSpan(characters.Slice(0, written), cache);
                        continue;
                    }
                }
#endif
                _materialized[ordinal] = ReuseFieldString(GetString(ordinal), cache);
            }
        }

        internal bool IsMissing(int ordinal)
        {
            ValidateOrdinal(ordinal);
            return _lengths[ordinal] < 0;
        }

        internal bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text)
        {
            ValidateOrdinal(ordinal);
            int length = _lengths[ordinal];
            if (length < 0)
            {
                text = default;
                return false;
            }

            ReadOnlySpan<byte> bytes = _buffer.AsSpan(_starts[ordinal], length);
            // A permissive decoder replaces malformed bytes. Such a raw segment would not
            // represent GetString's normalized text and cannot be borrowed as UTF-8.
            if (!_allFieldsAscii && !System.Text.Unicode.Utf8.IsValid(bytes))
            {
                text = default;
                return false;
            }

            text = bytes;
            return true;
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
