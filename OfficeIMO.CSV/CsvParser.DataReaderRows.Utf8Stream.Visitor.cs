#if NET8_0_OR_GREATER
using System.Buffers;
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
        private byte[] _buffer;
        private readonly Encoding _encoding;
        private static readonly string[] SingleCharacterStrings = CreateSingleCharacterStrings();
        private Field[] _fields;

        /// <summary>
        /// One field of the current row. <see cref="Start"/> is the byte offset of the raw text,
        /// <see cref="MissingStart"/> for a column the row does not have, or
        /// <see cref="MaterializedStart"/> when only <see cref="Materialized"/> holds the value.
        /// A negative <see cref="Length"/> is the bitwise complement of a quoted field's inner
        /// length when that text contains escaped quotes. A non-negative range that starts with a
        /// quote is a whole quoted field from the block scanner; its quotes are removed on access.
        /// </summary>
        private struct Field
        {
            internal int Start;
            internal int Length;
            internal string? Materialized;
        }

        private const int MissingStart = -1;
        private const int MaterializedStart = -2;
        private bool _allFieldsAscii;
        private bool _mayHaveEscapedQuotes;
        private char[]? _spanScratch;

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

        internal void SetBuffer(byte[] buffer) => _buffer = buffer;

        /// <summary>Stores a quoted field's inner text, deferring removal of escaped quotes.</summary>
        internal void VisitQuotedFieldRange(int fieldIndex, int start, int length, bool escaped) =>
            VisitFieldRange(fieldIndex, start, escaped ? ~length : length);

        /// <summary>Stores a field that was already unescaped and decoded.</summary>
        internal void VisitMaterializedField(int fieldIndex, string value)
        {
            VisitFieldRange(fieldIndex, MaterializedStart, 0);
            _fields[fieldIndex].Materialized = value;
        }

        internal readonly bool IsEmpty(int ordinal)
        {
            ref Field field = ref _fields[ordinal];
            if (field.Materialized is { } materialized)
            {
                return materialized.Length == 0;
            }

            return !TryResolveText(ordinal, out _, out int length, out _) || length == 0;
        }

        /// <summary>
        /// Resolves a field's text bytes without its quotes. Returns false for a missing field or one
        /// held only as a decoded string; <paramref name="escaped"/> reports escaped quote pairs.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private readonly bool TryResolveText(int ordinal, out int start, out int length, out bool escaped)
        {
            ref Field field = ref _fields[ordinal];
            start = field.Start;
            length = field.Length;
            escaped = false;
            if (start < 0)
            {
                return false;
            }

            if (length < 0)
            {
                length = ~length;
                escaped = true;
            }
            else if (length >= 2 && _buffer[start] == (byte)'"')
            {
                start++;
                length -= 2;
                escaped = _mayHaveEscapedQuotes && _buffer.AsSpan(start, length).Contains((byte)'"');
            }

            return true;
        }

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

        /// <summary>Publishes the fields in one 64-byte block after reserving their complete span.</summary>
        internal void VisitDelimiterMask(
            ulong delimiterMask,
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


        internal void Complete(
            int fieldCount,
            CsvColumnCountMismatchPolicy mismatchPolicy,
            bool allFieldsAscii,
            bool mayHaveEscapedQuotes)
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
            _mayHaveEscapedQuotes = mayHaveEscapedQuotes;

            int expectedCount = SourceColumnCount > 0 ? SourceColumnCount : fieldCount;
            for (int index = fieldCount; index < expectedCount; index++)
            {
                _fields[index].Start = MissingStart;
                _fields[index].Length = 0;
                _fields[index].Materialized = null;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal readonly string GetString(int ordinal)
        {
            ValidateOrdinal(ordinal);
            ref Field field = ref _fields[ordinal];
            // Raw text becomes a new string on every call. Only interned and already decoded values
            // are stored, so ordinary reads write no string references back into the field table.
            if (field.Materialized is { } materialized)
            {
                return materialized;
            }

            int length = field.Length;
            if (length > 0)
            {
                int start = field.Start;
                byte first = _buffer[start];
                if (first == (byte)'"' && length >= 2)
                {
                    return GetQuotedString(start + 1, length - 2);
                }

                if (length == 1 && first < SingleCharacterStrings.Length)
                {
                    return SingleCharacterStrings[first];
                }

                return _allFieldsAscii ? CreateAsciiString(_buffer, start, length) : DecodeString(start, length);
            }

            return length == 0 ? string.Empty : CreateUnescapedString(field.Start, ~length);
        }

        private readonly string GetQuotedString(int start, int length)
        {
            if (length == 0)
            {
                return string.Empty;
            }

            if (_mayHaveEscapedQuotes && _buffer.AsSpan(start, length).Contains((byte)'"'))
            {
                return CreateUnescapedString(start, length);
            }

            if (length == 1 && _buffer[start] < SingleCharacterStrings.Length)
            {
                return SingleCharacterStrings[_buffer[start]];
            }

            return _allFieldsAscii ? CreateAsciiString(_buffer, start, length) : DecodeString(start, length);
        }

        /// <summary>Compares a field with text without creating a string for ASCII raw text.</summary>
        internal readonly bool TextEquals(int ordinal, string value)
        {
            if (_fields[ordinal].Materialized is null && TryResolveText(ordinal, out int start, out int length, out bool escaped) && !escaped)
            {
                ReadOnlySpan<byte> bytes = _buffer.AsSpan(start, length);
                if (_allFieldsAscii || Ascii.IsValid(bytes))
                {
                    return Ascii.Equals(bytes, value);
                }
            }

            return string.Equals(GetString(ordinal), value, StringComparison.Ordinal);
        }

        /// <summary>
        /// Returns a field's text for immediate conversion. ASCII text is widened into a reused
        /// buffer, so the span is valid only until the next call; other text uses its string.
        /// </summary>
        internal ReadOnlySpan<char> GetSpan(int ordinal)
        {
            ValidateOrdinal(ordinal);
            if (!TryResolveText(ordinal, out int start, out int length, out bool escaped))
            {
                return (uint)ordinal < (uint)FieldCount ? GetString(ordinal).AsSpan() : default;
            }

            ReadOnlySpan<byte> bytes = _buffer.AsSpan(start, length);
            if (escaped || _fields[ordinal].Materialized is not null || !(_allFieldsAscii || Ascii.IsValid(bytes)))
            {
                return GetString(ordinal).AsSpan();
            }

            if (_spanScratch is null || _spanScratch.Length < length)
            {
                _spanScratch = new char[Math.Max(length, 256)];
            }

            Span<char> characters = _spanScratch.AsSpan(0, length);
            Encoding.Latin1.GetChars(bytes, characters);
            return characters;
        }

        private readonly string DecodeString(int start, int length)
        {
            ReadOnlySpan<byte> bytes = _buffer.AsSpan(start, length);
            return Ascii.IsValid(bytes) ? CreateAsciiString(_buffer, start, length) : _encoding.GetString(bytes);
        }

        /// <summary>Decodes quoted text whose escaped quote pairs each become one quote.</summary>
        private readonly string CreateUnescapedString(int start, int length)
        {
            ReadOnlySpan<byte> raw = _buffer.AsSpan(start, length);
            byte[]? rented = null;
            Span<byte> bytes = length <= 256
                ? stackalloc byte[256]
                : (rented = ArrayPool<byte>.Shared.Rent(length));
            int written = 0;
            while (true)
            {
                // Inside a scanned quoted field every quote is the first of an escaped pair.
                int quote = raw.IndexOf((byte)'"');
                if (quote < 0)
                {
                    raw.CopyTo(bytes.Slice(written));
                    written += raw.Length;
                    break;
                }

                raw.Slice(0, quote + 1).CopyTo(bytes.Slice(written));
                written += quote + 1;
                raw = raw.Slice(quote + 2);
            }

            ReadOnlySpan<byte> unescaped = bytes.Slice(0, written);
            string value = _allFieldsAscii || Ascii.IsValid(unescaped)
                ? Encoding.Latin1.GetString(unescaped)
                : _encoding.GetString(unescaped);
            if (rented is not null)
            {
                ArrayPool<byte>.Shared.Return(rented);
            }

            return value;
        }

        // Only validated ASCII reaches this method. A string allocated with '\0' is not filled
        // before it is returned to us, and the first and last widening steps overlap so every
        // load stays inside the source span and every store inside the new string.
        private static string CreateAsciiString(byte[] buffer, int start, int length)
        {
            if (!BitConverter.IsLittleEndian || !Vector128.IsHardwareAccelerated || length < 2)
            {
                return Encoding.Latin1.GetString(buffer.AsSpan(start, length));
            }

            string value = new string('\0', length);
            ref byte source = ref buffer[start];
            ref byte destination = ref Unsafe.As<char, byte>(ref MemoryMarshal.GetReference(value.AsSpan()));
            if (length < 16)
            {
                WidenShortAscii(ref source, ref destination, length);
            }
            else if (length < 32 || !Vector256.IsHardwareAccelerated)
            {
                for (int offset = 0; offset < length - 16; offset += 16)
                {
                    WidenSixteen(ref Unsafe.Add(ref source, offset), ref Unsafe.Add(ref destination, offset * 2));
                }

                WidenSixteen(ref Unsafe.Add(ref source, length - 16), ref Unsafe.Add(ref destination, (length - 16) * 2));
            }
            else
            {
                for (int offset = 0; offset < length - 32; offset += 32)
                {
                    WidenThirtyTwo(ref Unsafe.Add(ref source, offset), ref Unsafe.Add(ref destination, offset * 2));
                }

                WidenThirtyTwo(ref Unsafe.Add(ref source, length - 32), ref Unsafe.Add(ref destination, (length - 32) * 2));
            }

            return value;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static void WidenSixteen(ref byte source, ref byte destination)
        {
            (Vector128<ushort> lower, Vector128<ushort> upper) = Vector128.Widen(Vector128.LoadUnsafe(ref source));
            lower.AsByte().StoreUnsafe(ref destination);
            upper.AsByte().StoreUnsafe(ref destination, 16);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static void WidenThirtyTwo(ref byte source, ref byte destination)
        {
            (Vector256<ushort> lower, Vector256<ushort> upper) = Vector256.Widen(Vector256.LoadUnsafe(ref source));
            lower.AsByte().StoreUnsafe(ref destination);
            upper.AsByte().StoreUnsafe(ref destination, 32);
        }

        /// <summary>Widens an already validated 2..15-byte ASCII span within its exact destination length.</summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static void WidenShortAscii(ref byte source, ref byte destination, int length)
        {
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
                if (TryResolveText(ordinal, out int start, out int length, out bool escaped) && !escaped &&
                    length > 0 && length <= characters.Length)
                {
                    ReadOnlySpan<byte> bytes = _buffer.AsSpan(start, length);
                    if (_allFieldsAscii || Ascii.IsValid(bytes))
                    {
                        if (length == 1)
                        {
                            characters[0] = (char)bytes[0];
                        }
                        else if (length < 16 && BitConverter.IsLittleEndian && Vector128.IsHardwareAccelerated)
                        {
                            ref byte source = ref MemoryMarshal.GetReference(bytes);
                            ref byte destination = ref Unsafe.As<char, byte>(ref MemoryMarshal.GetReference(characters));
                            WidenShortAscii(ref source, ref destination, length);
                        }
                        else
                        {
                            Encoding.Latin1.GetChars(bytes, characters);
                        }
                        _fields[ordinal].Materialized = ReuseFieldSpan(characters.Slice(0, length), cache, lookup);
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
            return _fields[ordinal].Start == MissingStart;
        }

        internal bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text)
        {
            ValidateOrdinal(ordinal);
            // Missing, decoded and escaped fields have no borrowable raw UTF-8 text.
            if (!TryResolveText(ordinal, out int start, out int length, out bool escaped) || escaped)
            {
                text = default;
                return false;
            }

            ReadOnlySpan<byte> bytes = _buffer.AsSpan(start, length);
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
        private readonly void ValidateOrdinal(int ordinal)
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
