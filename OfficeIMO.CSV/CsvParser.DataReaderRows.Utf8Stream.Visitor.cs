using System.Buffers;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
#if NET8_0_OR_GREATER
using System.Numerics;
using System.Runtime.Intrinsics;
#endif

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
        /// <summary>
        /// When set, the current record's fields are read from a batch index of separator offsets
        /// instead of <see cref="_fields"/>: field <c>i</c> spans from after entry
        /// <c>_indexFirst + i - 1</c> to entry <c>_indexFirst + i</c>.
        /// </summary>
        private int[]? _indexEnds;
        private int _indexFirst;
        private bool _indexed;
        private bool _mayHaveEscapedQuotes;
        // Quote-free dialects store every field as plain text, including one that starts with a quote.
        private readonly bool _quotesAreData;
        private char[]? _spanScratch;

        internal CsvUtf8DataReaderRowVisitor(byte[] buffer, Encoding encoding, bool quotesAreData)
        {
            _buffer = buffer;
            _encoding = encoding;
            _quotesAreData = quotesAreData;
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

        internal void Reset()
        {
            FieldCount = 0;
            _indexed = false;
        }

        /// <summary>Sets the batch index that indexed records read; it is kept for the reader's lifetime.</summary>
        internal void SetBatchIndex(int[] ends) => _indexEnds = ends;

        /// <summary>Publishes a record whose fields end at <paramref name="count"/> consecutive batch index entries.</summary>
        internal void SetIndexedRecord(int first, int count)
        {
            _indexed = true;
            _indexFirst = first;
            FieldCount = count;
        }

        /// <summary>Reads a field's stored range; an indexed record has no decoded values.</summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private readonly void GetRange(int ordinal, out int start, out int length, out string? materialized)
        {
#if NET8_0_OR_GREATER
            if (_indexed)
            {
                int[] ends = _indexEnds!;
                // Indexed records are complete: every valid ordinal is below the record's field count.
                ref int end = ref Unsafe.Add(ref MemoryMarshal.GetArrayDataReference(ends), _indexFirst + ordinal);
                start = Unsafe.Add(ref end, -1) + 1;
                length = end - start;
                materialized = null;
                return;
            }
#endif

            ref Field field = ref _fields[ordinal];
            start = field.Start;
            length = field.Length;
            materialized = field.Materialized;
        }

        /// <summary>Copies an indexed record into the field table before values are added to it.</summary>
        private void MaterializeIndexedFields()
        {
            int count = FieldCount;
            EnsureCapacity(count);
            for (int ordinal = 0; ordinal < count; ordinal++)
            {
                GetRange(ordinal, out int start, out int length, out _);
                ref Field field = ref _fields[ordinal];
                field.Start = start;
                field.Length = length;
                field.Materialized = null;
            }

            _indexed = false;
        }

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
            GetRange(ordinal, out _, out _, out string? materialized);
            if (materialized is not null)
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
            GetRange(ordinal, out start, out length, out _);
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
            else if (length >= 2 && _buffer[start] == (byte)'"' && !_quotesAreData)
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

#if NET8_0_OR_GREATER
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
#endif

        internal void Complete(
            int fieldCount,
            CsvColumnCountMismatchPolicy mismatchPolicy,
            bool allFieldsAscii,
            bool mayHaveEscapedQuotes)
        {
            if (!_indexed)
            {
                EnsureCapacity(fieldCount);
            }

            FieldCount = fieldCount;
            int sourceColumnCount = SourceColumnCount;
            if (sourceColumnCount > 0 &&
                mismatchPolicy == CsvColumnCountMismatchPolicy.Strict &&
                fieldCount != sourceColumnCount)
            {
                ThrowColumnCountMismatch(fieldCount, sourceColumnCount);
            }

            _allFieldsAscii = allFieldsAscii;
            _mayHaveEscapedQuotes = mayHaveEscapedQuotes;
            if (sourceColumnCount > fieldCount)
            {
                MarkMissingFields(fieldCount, sourceColumnCount);
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ThrowColumnCountMismatch(int fieldCount, int sourceColumnCount) =>
            throw new CsvException($"Row contains {fieldCount} values but header defines {sourceColumnCount} columns.");

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void MarkMissingFields(int fieldCount, int expectedCount)
        {
            if (_indexed)
            {
                MaterializeIndexedFields();
            }

            for (int index = fieldCount; index < expectedCount; index++)
            {
                _fields[index].Start = MissingStart;
                _fields[index].Length = 0;
                _fields[index].Materialized = null;
            }
        }

        /// <summary>
        /// Finds the first of a record's fields whose text is longer than <paramref name="maxLength"/>
        /// characters, before <see cref="Complete"/> pads or limits the record to the header.
        /// </summary>
        internal int FindFieldLongerThan(int fieldCount, int maxLength, bool allFieldsAscii, bool mayHaveEscapedQuotes, out int length)
        {
            _allFieldsAscii = allFieldsAscii;
            _mayHaveEscapedQuotes = mayHaveEscapedQuotes;
            for (int ordinal = 0; ordinal < fieldCount; ordinal++)
            {
                ref Field field = ref _fields[ordinal];
                // UTF-8 never takes fewer bytes than UTF-16 code units, so the raw byte length
                // bounds the text length; only a field that may be too long is decoded.
                int bound = field.Materialized?.Length ?? (field.Length < 0 ? ~field.Length : field.Length);
                if (bound <= maxLength)
                {
                    continue;
                }

                length = field.Materialized?.Length ?? GetFieldString(ordinal).Length;
                if (length > maxLength)
                {
                    return ordinal;
                }
            }

            length = 0;
            return -1;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal readonly string GetString(int ordinal)
        {
            ValidateOrdinal(ordinal);
            return GetFieldString(ordinal);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private readonly string GetFieldString(int ordinal)
        {
            GetRange(ordinal, out int start, out int length, out string? materialized);
            // Raw text becomes a new string on every call. Only interned and already decoded values
            // are stored, so ordinary reads write no string references back into the field table.
            if (materialized is not null)
            {
                return materialized;
            }

            if (length > 0)
            {
                byte first = _buffer[start];
                if (first == (byte)'"' && length >= 2 && !_quotesAreData)
                {
                    return GetQuotedString(start + 1, length - 2);
                }

                if (length == 1 && first < SingleCharacterStrings.Length)
                {
                    return SingleCharacterStrings[first];
                }

                return _allFieldsAscii ? CreateAsciiString(_buffer, start, length) : DecodeString(start, length);
            }

            return length == 0 ? string.Empty : CreateUnescapedString(start, ~length);
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
            if (!HasMaterializedValue(ordinal) && TryResolveText(ordinal, out int start, out int length, out bool escaped) && !escaped)
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
            if (escaped || HasMaterializedValue(ordinal) || !(_allFieldsAscii || Ascii.IsValid(bytes)))
            {
                return GetString(ordinal).AsSpan();
            }

            if (_spanScratch is null || _spanScratch.Length < length)
            {
                _spanScratch = new char[Math.Max(length, 256)];
            }

            Span<char> characters = _spanScratch.AsSpan(0, length);
            CsvAsciiText.Widen(bytes, characters);
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
                ? CsvAsciiText.GetString(unescaped)
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
#if !NET8_0_OR_GREATER
            return CsvAsciiText.GetString(buffer, start, length);
#else
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
#endif
        }

#if NET8_0_OR_GREATER

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
#endif

        /// <summary>Completes requested per-read string sharing before exposing the row.</summary>
        internal void InternFields(CsvStringCache? cache)
        {
            if (cache is null) return;
            if (_indexed) MaterializeIndexedFields();
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
            return !_indexed && _fields[ordinal].Start == MissingStart;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private readonly bool HasMaterializedValue(int ordinal) => !_indexed && _fields[ordinal].Materialized is not null;

#if NET8_0_OR_GREATER
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
#endif

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

/// <summary>Turns validated ASCII bytes into characters, one byte per character.</summary>
internal static class CsvAsciiText
{
#if NET8_0_OR_GREATER
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static string GetString(ReadOnlySpan<byte> ascii) => Encoding.Latin1.GetString(ascii);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void Widen(ReadOnlySpan<byte> ascii, Span<char> characters) => Encoding.Latin1.GetChars(ascii, characters);
#else
    internal static string GetString(ReadOnlySpan<byte> ascii)
    {
        if (ascii.IsEmpty)
        {
            return string.Empty;
        }

        string value = new string('\0', ascii.Length);
        Widen(ref MemoryMarshal.GetReference(ascii), ref MemoryMarshal.GetReference(value.AsSpan()), ascii.Length);
        return value;
    }

    internal static string GetString(byte[] buffer, int start, int length)
    {
        if (length == 0)
        {
            return string.Empty;
        }

        // A new string is not shared until it is returned, so it can be filled in place.
        string value = new string('\0', length);
        Widen(ref buffer[start], ref MemoryMarshal.GetReference(value.AsSpan()), length);
        return value;
    }

    /// <summary>Widens four bytes at a time into a word of four characters.</summary>
    private static void Widen(ref byte source, ref char destination, int length)
    {
        int index = 0;
        if (BitConverter.IsLittleEndian)
        {
            for (; index <= length - 4; index += 4)
            {
                ulong word = Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref source, index));
                word = (word | (word << 16)) & 0x0000ffff0000ffffUL;
                word = (word | (word << 8)) & 0x00ff00ff00ff00ffUL;
                Unsafe.WriteUnaligned(ref Unsafe.As<char, byte>(ref Unsafe.Add(ref destination, index)), word);
            }
        }

        for (; index < length; index++)
        {
            Unsafe.Add(ref destination, index) = (char)Unsafe.Add(ref source, index);
        }
    }

    internal static void Widen(ReadOnlySpan<byte> ascii, Span<char> characters)
    {
        for (int index = 0; index < ascii.Length; index++)
        {
            characters[index] = (char)ascii[index];
        }
    }
#endif
}
