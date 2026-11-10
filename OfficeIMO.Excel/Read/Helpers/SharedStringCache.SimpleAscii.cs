using System.Text;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif

namespace OfficeIMO.Excel {
    internal sealed partial class SharedStringCache {
#if NET8_0_OR_GREATER
        private const int SimpleAsciiScanChunkBytes = 16 * 1024;
        private static readonly SearchValues<byte> SimpleAsciiTextSpecialBytes = SearchValues.Create(
            new byte[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15,
                16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31,
                0x7f, (byte)'<', (byte)'_', (byte)']' });
#endif

        /// <summary>
        /// Validates the complete plain ASCII SpreadsheetML shape before retaining text offsets.
        /// Markup, escaping, Unicode, or different attribute layouts use the full XML path.
        /// </summary>
        private bool TryIndexSimpleAsciiItems(OpenXmlPooledPartStream stream, out IndexedAsciiItems? items) {
            items = null;
            byte[] bytes = stream.BorrowBuffer(out int length);
            if (!TryReadSimpleAsciiHeader(bytes, length, out int position, out int declaredUniqueCount,
                    out int prefixStart, out int prefixLength)) {
                return false;
            }

            // Reject common full-XML cases before allocating any item strings.
            // A late Unicode value or entity otherwise discards the ASCII prefix
            // and allocates it again when the XML reader reparses the whole part.
#if NET8_0_OR_GREATER
            for (int index = position; index < length;) {
                _cancellationToken.ThrowIfCancellationRequested();
                int chunkLength = Math.Min(SimpleAsciiScanChunkBytes, length - index);
                ReadOnlySpan<byte> content = bytes.AsSpan(index, chunkLength);
                if (!Ascii.IsValid(content) || content.IndexOf((byte)'&') >= 0) return false;
                index += chunkLength;
            }
#else
            for (int index = position; index < length; index++) {
                if (bytes[index] >= 0x80 || bytes[index] == (byte)'&') return false;
            }
#endif

            var parsed = IndexedAsciiEntryPool.Rent(Math.Min(declaredUniqueCount,
                Math.Min(_maxSharedStringItems, length / 7)));
            bool transferred = false;
            try {
                long totalCharacters = 0;
                while (!ConsumeQualified(bytes, length, ref position, "</", prefixStart, prefixLength, "sst>")) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    bool preservesWhitespace = false;
                    if (!ConsumeQualified(bytes, length, ref position, "<", prefixStart, prefixLength, "si>")
                        || !ConsumeQualified(bytes, length, ref position, "<", prefixStart, prefixLength, "t")) {
                        return false;
                    }
                    if (!Consume(bytes, length, ref position, ">")) {
                        if (!Consume(bytes, length, ref position, " xml:space=\"preserve\">")) {
                            return false;
                        }
                        preservesWhitespace = true;
                    }

                    int textStart = position;
#if NET8_0_OR_GREATER
                    if (!TryScanSimpleAsciiText(bytes, length, ref position, preservesWhitespace, out bool onlyWhitespace)) {
                        return false;
                    }
#else
                    while (position < length && bytes[position] != (byte)'<') {
                        if ((position & 16383) == 0) _cancellationToken.ThrowIfCancellationRequested();
                        byte value = bytes[position];
                        if (value < 0x20 || value > 0x7E || value == (byte)'&'
                            || value == (byte)']' && position + 2 < length
                                && bytes[position + 1] == (byte)']'
                                && bytes[position + 2] == (byte)'>') {
                            return false;
                        }
                        // Leave SpreadsheetML escaped text to the canonical XML/SDK loader.
                        if (value == (byte)'_' && position + 1 < length
                            && bytes[position + 1] is (byte)'x' or (byte)'X') return false;
                        position++;
                    }
#endif

                    int textLength = position - textStart;
                    if (!preservesWhitespace && textLength > 0) {
#if !NET8_0_OR_GREATER
                        bool onlyWhitespace = true;
                        for (int index = textStart; index < position; index++) {
                            if (bytes[index] != (byte)' ') {
                                onlyWhitespace = false;
                                break;
                            }
                        }
#endif
                        if (onlyWhitespace) return false;
                    }
                    if (!ConsumeQualified(bytes, length, ref position, "</", prefixStart, prefixLength, "t>")
                        || !ConsumeQualified(bytes, length, ref position, "</", prefixStart, prefixLength, "si>")) {
                        return false;
                    }

                    EnsureCanAddSharedString(parsed.Count);
                    ValidateSharedStringLength(textLength, ref totalCharacters);
                    parsed.Add(new AsciiTextEntry(textStart, textLength));
                }

                SkipWhitespace(bytes, length, ref position);
                if (position != length) {
                    return false;
                }

                _cancellationToken.ThrowIfCancellationRequested();
                items = new IndexedAsciiItems(stream, bytes, parsed);
                transferred = true;
                return true;
            } finally {
                if (!transferred) IndexedAsciiEntryPool.Return(parsed, retain: false);
            }
        }

#if NET8_0_OR_GREATER
        /// <summary>Skips ordinary text in bounded chunks without changing the scanner's XML fallback rules.</summary>
        private bool TryScanSimpleAsciiText(
            byte[] bytes,
            int length,
            ref int position,
            bool preservesWhitespace,
            out bool onlyWhitespace) {
            onlyWhitespace = true;
            int chunkEnd = position + Math.Min(SimpleAsciiScanChunkBytes, length - position);
            while (position < length) {
                if (position == chunkEnd) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    chunkEnd = position + Math.Min(SimpleAsciiScanChunkBytes, length - position);
                }
                ReadOnlySpan<byte> chunk = bytes.AsSpan(position, chunkEnd - position);
                int special = chunk.IndexOfAny(SimpleAsciiTextSpecialBytes);
                int textLength = special < 0 ? chunk.Length : special;
                if (!preservesWhitespace && onlyWhitespace
                    && chunk.Slice(0, textLength).IndexOfAnyExcept((byte)' ') >= 0) {
                    onlyWhitespace = false;
                }
                position += textLength;
                if (special < 0) continue;

                byte value = bytes[position];
                if (value == (byte)'<') return true;
                if (value == (byte)'_') {
                    if (position + 1 < length && bytes[position + 1] is (byte)'x' or (byte)'X') return false;
                } else if (value == (byte)']') {
                    if (position + 2 < length && bytes[position + 1] == (byte)']' && bytes[position + 2] == (byte)'>') return false;
                } else {
                    return false;
                }

                onlyWhitespace = false;
                position++;
            }
            return true;
        }
#endif

        /// <summary>Limits large-part buffering to the simple table header handled by this scanner.</summary>
        internal static bool CanIndexSimpleAsciiPrefix(byte[] bytes, int length) =>
            TryReadSimpleAsciiHeader(bytes, length, out _, out _, out _, out _);

        /// <summary>
        /// Accepts the declarations Excel and the Open XML SDK write: an unprefixed table bound by
        /// the default namespace, or one prefix bound on the root and used by every element.
        /// </summary>
        private static bool TryReadSimpleAsciiHeader(
            byte[] bytes,
            int length,
            out int position,
            out int declaredUniqueCount,
            out int prefixStart,
            out int prefixLength) {
            position = 0;
            declaredUniqueCount = 0;
            prefixStart = prefixLength = 0;
            if (!Consume(bytes, length, ref position, "<?xml version=\"1.0\" encoding=\"")
                || !ConsumeIgnoreCase(bytes, length, ref position, "utf-8")
                || !Consume(bytes, length, ref position, "\"")) return false;
            Consume(bytes, length, ref position, " standalone=\"yes\"");
            if (!Consume(bytes, length, ref position, "?>")) return false;
            SkipWhitespace(bytes, length, ref position);
            if (!Consume(bytes, length, ref position, "<")) return false;

            if (!Consume(bytes, length, ref position, "sst ")) {
                prefixStart = position;
                if (position >= length || !IsAsciiNameStart(bytes[position])) return false;
                while (position < length && IsAsciiNameCharacter(bytes[position])) position++;
                if (!Consume(bytes, length, ref position, ":")) return false;
                prefixLength = position - prefixStart;
                if (!Consume(bytes, length, ref position, "sst xmlns:")
                    || !ConsumeBytes(bytes, length, ref position, prefixStart, prefixLength - 1)
                    || !Consume(bytes, length, ref position, "=\"")) return false;
            } else if (!Consume(bytes, length, ref position, "xmlns=\"")) {
                return false;
            }

            if (!Consume(bytes, length, ref position, "http://schemas.openxmlformats.org/spreadsheetml/2006/main\"")) return false;
            if (Consume(bytes, length, ref position, " count=\"")
                && (!ReadUnsignedDecimal(bytes, length, ref position, out _) || !Consume(bytes, length, ref position, "\""))) return false;
            if (Consume(bytes, length, ref position, " uniqueCount=\"")
                && (!ReadUnsignedDecimal(bytes, length, ref position, out declaredUniqueCount) || !Consume(bytes, length, ref position, "\""))) return false;
            if (!Consume(bytes, length, ref position, ">")) return false;
            if (declaredUniqueCount == 0) declaredUniqueCount = CountItemTags(bytes, length, position, prefixStart, prefixLength);
            return true;
        }

        /// <summary>
        /// Sizes the index of a table without a uniqueCount, so it is neither grown nor oversized.
        /// A miscount only changes the initial capacity; the scan still validates every item.
        /// </summary>
        private static int CountItemTags(byte[] bytes, int length, int position, int prefixStart, int prefixLength) {
            Span<byte> tag = stackalloc byte[prefixLength + 4];
            tag[0] = (byte)'<';
            bytes.AsSpan(prefixStart, prefixLength).CopyTo(tag.Slice(1));
            tag[prefixLength + 1] = (byte)'s';
            tag[prefixLength + 2] = (byte)'i';
            tag[prefixLength + 3] = (byte)'>';
            ReadOnlySpan<byte> remaining = bytes.AsSpan(position, length - position);
            int count = 0;
            for (int found; (found = remaining.IndexOf(tag)) >= 0; remaining = remaining.Slice(found + tag.Length)) {
                count++;
            }
            return count;
        }

        private static bool ConsumeQualified(byte[] bytes, int length, ref int position, string lead, int prefixStart, int prefixLength, string localName) {
            int start = position;
            if (Consume(bytes, length, ref position, lead)
                && ConsumeBytes(bytes, length, ref position, prefixStart, prefixLength)
                && Consume(bytes, length, ref position, localName)) return true;
            position = start;
            return false;
        }

        private static bool ConsumeBytes(byte[] bytes, int length, ref int position, int start, int count) {
            if (count > length - position || !bytes.AsSpan(position, count).SequenceEqual(bytes.AsSpan(start, count))) return false;
            position += count;
            return true;
        }

        private static bool ConsumeIgnoreCase(byte[] bytes, int length, ref int position, string value) {
            if (value.Length > length - position) return false;
            for (int index = 0; index < value.Length; index++) {
                int current = bytes[position + index];
                if (current is >= 'A' and <= 'Z') current += 'a' - 'A';
                if (current != value[index]) return false;
            }
            position += value.Length;
            return true;
        }

        private static bool IsAsciiNameStart(byte value) =>
            value is >= (byte)'A' and <= (byte)'Z' or >= (byte)'a' and <= (byte)'z' or (byte)'_';

        private static bool IsAsciiNameCharacter(byte value) =>
            IsAsciiNameStart(value) || value is >= (byte)'0' and <= (byte)'9' or (byte)'.' or (byte)'-';

        private static bool Consume(byte[] bytes, int length, ref int position, string value) {
            if (value.Length > length - position) return false;
#if NET8_0_OR_GREATER
            if (!Ascii.Equals(bytes.AsSpan(position, value.Length), value.AsSpan())) return false;
#else
            for (int index = 0; index < value.Length; index++) {
                if (bytes[position + index] != (byte)value[index]) return false;
            }
#endif
            position += value.Length;
            return true;
        }

        private static bool ReadUnsignedDecimal(byte[] bytes, int length, ref int position, out int value) {
            value = 0;
            int start = position;
            while (position < length && bytes[position] >= (byte)'0' && bytes[position] <= (byte)'9') {
                int digit = bytes[position++] - (byte)'0';
                if (value > (int.MaxValue - digit) / 10) return false;
                value = value * 10 + digit;
            }
            return position > start;
        }

        private static void SkipWhitespace(byte[] bytes, int length, ref int position) {
            while (position < length && bytes[position] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n') {
                position++;
            }
        }
    }
}
