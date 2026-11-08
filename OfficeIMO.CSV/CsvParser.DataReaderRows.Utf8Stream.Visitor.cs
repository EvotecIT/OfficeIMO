#if NET8_0_OR_GREATER
using System.Numerics;
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
        private Field[] _fields;

        private struct Field
        {
            internal int Start;
            internal int Length;
            internal string? Materialized;
        }
        private bool _allFieldsAscii;

        internal CsvUtf8DataReaderRowVisitor(byte[] buffer, Encoding encoding)
        {
            _buffer = buffer;
            _encoding = encoding;
            _fields = new Field[32];
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
            if ((uint)fieldIndex >= (uint)_fields.Length)
            {
                EnsureCapacity(fieldIndex + 1);
            }

            FieldCount = fieldIndex + 1;
            ref Field field = ref _fields[fieldIndex];
            field.Start = start;
            field.Length = length;
            field.Materialized = null;
        }

        /// <summary>Publishes the fields in one scanned block after reserving their complete span.</summary>
        internal void VisitDelimiterMask(
            uint delimiterMask,
            int chunkStart,
            ref int fieldStart,
            ref int fieldIndex)
        {
            if (delimiterMask == 0)
            {
                return;
            }

            int delimiterCount = BitOperations.PopCount(delimiterMask);
            int completedFieldCount = fieldIndex + delimiterCount;
            EnsureCapacity(completedFieldCount);
            Span<Field> fields = _fields.AsSpan(fieldIndex, delimiterCount);
            for (int index = 0; index < fields.Length; index++)
            {
                int delimiterIndex = chunkStart + BitOperations.TrailingZeroCount(delimiterMask);
                ref Field field = ref fields[index];
                field.Start = fieldStart;
                field.Length = delimiterIndex - fieldStart;
                field.Materialized = null;
                fieldStart = delimiterIndex + 1;
                delimiterMask &= delimiterMask - 1;
            }

            fieldIndex = completedFieldCount;
            FieldCount = completedFieldCount;
        }

        internal void ShiftStarts(int offset)
        {
            for (int index = 0; index < FieldCount; index++)
            {
                _fields[index].Start += offset;
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
                _fields[index].Start = -1;
                _fields[index].Length = -1;
                _fields[index].Materialized = null;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal string GetString(int ordinal)
        {
            ValidateOrdinal(ordinal);
            ref Field field = ref _fields[ordinal];
            int length = field.Length;
            if (length <= 0)
            {
                return string.Empty;
            }

            string? materialized = field.Materialized;
            if (materialized is null)
            {
                int start = field.Start;
                if (length == 1 && _buffer[start] < SingleCharacterStrings.Length)
                {
                    materialized = SingleCharacterStrings[_buffer[start]];
                }
                else
                {
                    if (_allFieldsAscii)
                    {
                        materialized = CreateAsciiString(_buffer, start, length);
                    }
                    else
                    {
                        ReadOnlySpan<byte> bytes = _buffer.AsSpan(start, length);
                        materialized = Ascii.IsValid(bytes)
                            ? CreateAsciiString(_buffer, start, length)
                            : _encoding.GetString(bytes);
                    }
                }

                field.Materialized = materialized;
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
                // The caller admits 2..15 bytes. Each endpoint load stays inside
                // that span, and its widened store ends within the exact string length.
                ref byte source = ref MemoryMarshal.GetReference(bytes);
                ref byte destination = ref Unsafe.As<char, byte>(ref MemoryMarshal.GetReference(characters));
                int length = bytes.Length;
                if (length >= 8)
                {
                    WidenEight(ref source, ref destination);
                    WidenEight(ref Unsafe.Add(ref source, length - 8), ref Unsafe.Add(ref destination, (length - 8) * 2));
                }
                else if (length >= 4)
                {
                    WidenFour(ref source, ref destination);
                    WidenFour(ref Unsafe.Add(ref source, length - 4), ref Unsafe.Add(ref destination, (length - 4) * 2));
                }
                else
                {
                    WidenTwo(ref source, ref destination);
                    WidenTwo(ref Unsafe.Add(ref source, length - 2), ref Unsafe.Add(ref destination, (length - 2) * 2));
                }
            });
        }

        private static void WidenEight(ref byte source, ref byte destination)
        {
            ulong word = Unsafe.ReadUnaligned<ulong>(ref source);
            Vector128.WidenLower(Vector128.CreateScalarUnsafe(word).AsByte()).AsByte().StoreUnsafe(ref destination);
        }

        private static void WidenFour(ref byte source, ref byte destination)
        {
            ulong word = Unsafe.ReadUnaligned<uint>(ref source);
            word = (word | (word << 16)) & 0x0000ffff0000ffffUL;
            word = (word | (word << 8)) & 0x00ff00ff00ff00ffUL;
            Unsafe.WriteUnaligned(ref destination, word);
        }

        private static void WidenTwo(ref byte source, ref byte destination)
        {
            uint word = Unsafe.ReadUnaligned<ushort>(ref source);
            Unsafe.WriteUnaligned(ref destination, (word | (word << 8)) & 0x00ff00ffU);
        }

        /// <summary>Completes requested per-read string sharing before exposing the row.</summary>
        internal void InternFields(CsvStringCache? cache)
        {
            if (cache is null) return;
            int count = SourceColumnCount > 0 ? Math.Min(FieldCount, SourceColumnCount) : FieldCount;
#if NET9_0_OR_GREATER
            var lookup = cache.GetAlternateLookup<ReadOnlySpan<char>>();
            Span<char> characters = stackalloc char[128];
#endif
            for (int ordinal = 0; ordinal < count; ordinal++)
            {
#if NET9_0_OR_GREATER
                int length = _fields[ordinal].Length;
                if (length > 0 && length <= characters.Length)
                {
                    ReadOnlySpan<byte> bytes = _buffer.AsSpan(_fields[ordinal].Start, length);
                    if (_allFieldsAscii || Ascii.IsValid(bytes))
                    {
                        int written = Encoding.Latin1.GetChars(bytes, characters);
                        _fields[ordinal].Materialized = ReuseFieldSpan(characters.Slice(0, written), cache, lookup);
                        continue;
                    }
                }
#endif
                _fields[ordinal].Materialized = ReuseFieldString(GetString(ordinal), cache);
            }
        }

        internal bool IsMissing(int ordinal)
        {
            ValidateOrdinal(ordinal);
            return _fields[ordinal].Length < 0;
        }

        internal bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text)
        {
            ValidateOrdinal(ordinal);
            int length = _fields[ordinal].Length;
            if (length < 0)
            {
                text = default;
                return false;
            }

            ReadOnlySpan<byte> bytes = _buffer.AsSpan(_fields[ordinal].Start, length);
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
            if (count <= _fields.Length)
            {
                return;
            }

            int capacity = _fields.Length;
            while (capacity < count)
            {
                capacity = checked(capacity * 2);
            }

            Array.Resize(ref _fields, capacity);
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
