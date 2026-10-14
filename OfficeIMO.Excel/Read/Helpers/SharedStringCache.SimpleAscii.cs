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
            if (!TryReadSimpleAsciiHeader(bytes, length, out int position, out int declaredUniqueCount)) {
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
                while (!Consume(bytes, length, ref position, "</sst>")) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    bool preservesWhitespace = false;
                    if (!Consume(bytes, length, ref position, "<si><t>")) {
                        if (!Consume(bytes, length, ref position, "<si><t xml:space=\"preserve\">")) {
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
                    if (!Consume(bytes, length, ref position, "</t></si>")) {
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
            TryReadSimpleAsciiHeader(bytes, length, out _, out _);

        private static bool TryReadSimpleAsciiHeader(byte[] bytes, int length, out int position, out int declaredUniqueCount) {
            position = 0;
            declaredUniqueCount = 0;
            if (!Consume(bytes, length, ref position, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")) return false;
            SkipWhitespace(bytes, length, ref position);
            return Consume(bytes, length, ref position,
                    "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"")
                && ReadUnsignedDecimal(bytes, length, ref position, out _)
                && Consume(bytes, length, ref position, "\" uniqueCount=\"")
                && ReadUnsignedDecimal(bytes, length, ref position, out declaredUniqueCount)
                && Consume(bytes, length, ref position, "\">");
        }

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
