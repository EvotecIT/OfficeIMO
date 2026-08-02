#nullable enable

using System.Text;
#if NET8_0_OR_GREATER
using System.Buffers;
using System.Text.Unicode;
#endif

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private const int MaximumFastXmlDepth = 64;
            private const int MaximumFastXmlAttributes = 32;
            private static readonly UTF8Encoding StrictUtf8 = new(
                encoderShouldEmitUTF8Identifier: false,
                throwOnInvalidBytes: true);
#if NET8_0_OR_GREATER
            private static readonly SearchValues<byte> InvalidXmlControlBytes = SearchValues.Create(
                new byte[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31 });
#endif

            /// <summary>
            /// Proves full well-formedness for the common, unprefixed SpreadsheetML shape.
            /// Richer XML constructs fall back to XmlReader validation rather than weakening it.
            /// </summary>
            private bool IsCanonicalWorksheetXmlFullyValidated() {
                if (!_sheetDataSupportsFastValidation) {
                    return false;
                }
#if NET8_0_OR_GREATER
                ReadOnlySpan<byte> document = _buffer!.AsSpan(0, _length);
                if (!Utf8.IsValid(document)
                    || document.IndexOf((byte)'&') >= 0
                    || document.IndexOfAny(InvalidXmlControlBytes) >= 0) {
                    return false;
                }
#else
                try {
                    _ = StrictUtf8.GetCharCount(_buffer!, 0, _length);
                } catch (DecoderFallbackException) {
                    return false;
                }

                for (int index = 0; index < _length; index++) {
                    byte value = _buffer![index];
                    if (value == (byte)'&'
                        || value < 0x20 && value is not (byte)'\t' and not (byte)'\r' and not (byte)'\n') {
                        return false;
                    }
                }
#endif

                if (_sheetDataContentStart >= 0
                    && _sheetDataEndTagStart >= _sheetDataContentStart) {
                    ReadOnlySpan<byte> sheetData = _buffer!.AsSpan(
                        _sheetDataContentStart,
                        _sheetDataEndTagStart - _sheetDataContentStart);
#if NET8_0_OR_GREATER
                    if (sheetData.IndexOf("<!"u8) >= 0 || sheetData.IndexOf("<?"u8) >= 0) {
                        return false;
                    }
#else
                    for (int index = 0; index + 1 < sheetData.Length; index++) {
                        if (sheetData[index] == (byte)'<'
                            && sheetData[index + 1] is (byte)'!' or (byte)'?') {
                            return false;
                        }
                    }
#endif
                }

                Span<int> nameStarts = stackalloc int[MaximumFastXmlDepth];
                Span<int> nameLengths = stackalloc int[MaximumFastXmlDepth];
                Span<int> namespacePrefixStarts = stackalloc int[MaximumFastXmlAttributes];
                Span<int> namespacePrefixLengths = stackalloc int[MaximumFastXmlAttributes];
                int depth = 0;
                int namespacePrefixCount = 0;
                int position = SkipUtf8PreambleAndWhitespace(0);
                bool rootSeen = false;
                bool rootClosed = false;
                if (!HasOnlyRootDefaultNamespace()) {
                    return false;
                }

                while (position < _length) {
                    if (_sheetDataContentStart >= 0
                        && position == _sheetDataContentStart
                        && _sheetDataEndTagStart >= position) {
                        position = _sheetDataEndTagStart;
                    }
                    int relative = _buffer!.AsSpan(position, _length - position).IndexOf((byte)'<');
                    if (relative < 0) {
                        return rootClosed && ContainsOnlyAsciiWhitespace(position, _length);
                    }
                    int tagStart = position + relative;
                    if (rootClosed && !ContainsOnlyAsciiWhitespace(position, tagStart)) {
                        return false;
                    }
                    if (tagStart + 1 >= _length) {
                        return false;
                    }

                    byte marker = _buffer![tagStart + 1];
                    if (marker == (byte)'?') {
                        if (rootSeen || !TrySkipCanonicalXmlDeclaration(tagStart, out position)) {
                            return false;
                        }
                        continue;
                    }
                    if (marker == (byte)'!'
                        || !TryParseTag(tagStart, _length, out Utf8Tag tag)
                        || !IsValidUnprefixedXmlName(tag.NameStart, tag.NameEnd)) {
                        return false;
                    }

                    if (!ValidateCanonicalTagAttributes(
                            tag,
                            out bool declaresDefaultNamespace,
                            out bool declaresSpreadsheetNamespace,
                            out _)) {
                        return false;
                    }
                    bool isRootStartTag = !rootSeen && !tag.IsEnd;
                    if (!ValidateCanonicalNamespaceUsage(
                            tag,
                            isRootStartTag,
                            namespacePrefixStarts,
                            namespacePrefixLengths,
                            ref namespacePrefixCount)) {
                        return false;
                    }

                    int nameLength = tag.NameEnd - tag.NameStart;
                    if (tag.IsEnd) {
                        if (tag.IsEmpty
                            || declaresDefaultNamespace
                            || depth == 0
                            || !ByteRangesEqual(
                                nameStarts[depth - 1],
                                nameLengths[depth - 1],
                                tag.NameStart,
                                nameLength)) {
                            return false;
                        }
                        depth--;
                        if (depth == 0) {
                            rootClosed = true;
                        }
                    } else {
                        if (!rootSeen) {
                            if (!AsciiEquals(tag.NameStart, nameLength, "worksheet")
                                || !declaresSpreadsheetNamespace) {
                                return false;
                            }
                            rootSeen = true;
                        } else if (rootClosed || declaresDefaultNamespace) {
                            return false;
                        }

                        if (!tag.IsEmpty) {
                            if (depth == MaximumFastXmlDepth) {
                                return false;
                            }
                            nameStarts[depth] = tag.NameStart;
                            nameLengths[depth] = nameLength;
                            depth++;
                        } else if (depth == 0) {
                            rootClosed = true;
                        }
                    }

                    position = tag.End + 1;
                }

                return rootSeen && rootClosed && depth == 0;
            }

            private bool TrySkipCanonicalXmlDeclaration(int start, out int nextPosition) {
                nextPosition = start;
                int preambleLength = _length >= 3
                    && _buffer![0] == 0xEF
                    && _buffer[1] == 0xBB
                    && _buffer[2] == 0xBF
                        ? 3
                        : 0;
                if (start != preambleLength
                    || start + 5 >= _length
                    || !AsciiEquals(start + 2, 3, "xml")
                    || !IsAsciiWhitespace(_buffer![start + 5])) {
                    return false;
                }

                int instructionEnd = IndexOfSequence(start + 5, _length, (byte)'?', (byte)'>');
                if (instructionEnd < 0) {
                    return false;
                }

                int position = start + 5;
                bool sawVersion = false;
                bool sawEncoding = false;
                bool sawStandalone = false;
                while (position < instructionEnd) {
                    int whitespaceStart = position;
                    while (position < instructionEnd && IsAsciiWhitespace(_buffer[position])) {
                        position++;
                    }
                    if (position == instructionEnd) {
                        break;
                    }
                    if (position == whitespaceStart) {
                        return false;
                    }

                    int nameStart = position;
                    while (position < instructionEnd && IsAsciiXmlNameCharacter(_buffer[position])) {
                        position++;
                    }
                    int nameLength = position - nameStart;
                    while (position < instructionEnd && IsAsciiWhitespace(_buffer[position])) {
                        position++;
                    }
                    if (nameLength == 0 || position >= instructionEnd || _buffer[position++] != (byte)'=') {
                        return false;
                    }
                    while (position < instructionEnd && IsAsciiWhitespace(_buffer[position])) {
                        position++;
                    }
                    if (position >= instructionEnd || _buffer[position] is not ((byte)'\"') and not ((byte)'\'')) {
                        return false;
                    }

                    byte quote = _buffer[position++];
                    int valueStart = position;
                    while (position < instructionEnd && _buffer[position] != quote) {
                        position++;
                    }
                    if (position >= instructionEnd) {
                        return false;
                    }
                    int valueLength = position - valueStart;
                    position++;

                    if (!sawVersion && AsciiEquals(nameStart, nameLength, "version")) {
                        if (!AsciiEquals(valueStart, valueLength, "1.0")) {
                            return false;
                        }
                        sawVersion = true;
                    } else if (sawVersion
                               && !sawEncoding
                               && !sawStandalone
                               && AsciiEquals(nameStart, nameLength, "encoding")) {
                        if (!IsValidXmlEncodingName(valueStart, valueLength)) {
                            return false;
                        }
                        sawEncoding = true;
                    } else if (sawVersion
                               && !sawStandalone
                               && AsciiEquals(nameStart, nameLength, "standalone")) {
                        if (!AsciiEquals(valueStart, valueLength, "yes")
                            && !AsciiEquals(valueStart, valueLength, "no")) {
                            return false;
                        }
                        sawStandalone = true;
                    } else {
                        return false;
                    }
                }

                if (!sawVersion) {
                    return false;
                }
                nextPosition = instructionEnd + 2;
                return true;
            }

            private bool IsValidXmlEncodingName(int start, int length) {
                if (length == 0 || !IsAsciiLetter(_buffer![start])) {
                    return false;
                }
                for (int index = start + 1; index < start + length; index++) {
                    byte value = _buffer![index];
                    if (!IsAsciiLetter(value)
                        && value is not (>= (byte)'0' and <= (byte)'9')
                        and not (byte)'.'
                        and not (byte)'_'
                        and not (byte)'-') {
                        return false;
                    }
                }
                return true;
            }

            private static bool IsAsciiLetter(byte value) =>
                value is >= (byte)'A' and <= (byte)'Z'
                or >= (byte)'a' and <= (byte)'z';

            private bool HasOnlyRootDefaultNamespace() {
                int searchStart = 0;
                int defaultNamespaceCount = 0;
                while (searchStart < _length) {
#if NET8_0_OR_GREATER
                    int relative = _buffer!.AsSpan(searchStart, _length - searchStart).IndexOf("xmlns"u8);
#else
                    int relative = _buffer!.AsSpan(searchStart, _length - searchStart).IndexOf((byte)'x');
#endif
                    if (relative < 0) {
                        return defaultNamespaceCount == 1;
                    }
                    int match = searchStart + relative;
#if !NET8_0_OR_GREATER
                    if (match + 5 > _length
                        || _buffer![match + 1] != (byte)'m'
                        || _buffer[match + 2] != (byte)'l'
                        || _buffer[match + 3] != (byte)'n'
                        || _buffer[match + 4] != (byte)'s') {
                        searchStart = match + 1;
                        continue;
                    }
#endif
                    int position = match + 5;
                    while (position < _length && IsAsciiWhitespace(_buffer![position])) {
                        position++;
                    }
                    if (position < _length && _buffer![position] == (byte)'=') {
                        defaultNamespaceCount++;
                        if (defaultNamespaceCount > 1) {
                            return false;
                        }
                    }
                    searchStart = position + 1;
                }
                return defaultNamespaceCount == 1;
            }

            private int SkipUtf8PreambleAndWhitespace(int position) {
                if (_length >= 3
                    && _buffer![0] == 0xEF
                    && _buffer[1] == 0xBB
                    && _buffer[2] == 0xBF) {
                    position = 3;
                }
                while (position < _length && IsAsciiWhitespace(_buffer![position])) {
                    position++;
                }
                return position;
            }

            private bool ValidateCanonicalTagAttributes(
                Utf8Tag tag,
                out bool declaresDefaultNamespace,
                out bool declaresSpreadsheetNamespace,
                out bool hasPrefixedAttributes) {
                declaresDefaultNamespace = false;
                declaresSpreadsheetNamespace = false;
                hasPrefixedAttributes = false;
                Span<int> starts = stackalloc int[MaximumFastXmlAttributes];
                Span<int> lengths = stackalloc int[MaximumFastXmlAttributes];
                int count = 0;
                int position = tag.NameEnd;
                while (position < tag.End) {
                    int whitespaceStart = position;
                    while (position < tag.End && IsAsciiWhitespace(_buffer![position])) {
                        position++;
                    }
                    if (position >= tag.End) {
                        return true;
                    }
                    if (_buffer![position] == (byte)'/') {
                        position++;
                        return !tag.IsEnd && position == tag.End;
                    }
                    if (position == whitespaceStart
                        || tag.IsEnd
                        || count == MaximumFastXmlAttributes) {
                        return false;
                    }

                    int nameStart = position;
                    while (position < tag.End && !IsAttributeNameTerminator(_buffer[position])) {
                        position++;
                    }
                    int nameLength = position - nameStart;
                    if (!IsValidXmlAttributeName(nameStart, position)) {
                        return false;
                    }
                    for (int index = nameStart; index < position; index++) {
                        if (_buffer![index] == (byte)':') {
                            hasPrefixedAttributes = true;
                            break;
                        }
                    }
                    for (int index = 0; index < count; index++) {
                        if (ByteRangesEqual(starts[index], lengths[index], nameStart, nameLength)) {
                            return false;
                        }
                    }
                    starts[count] = nameStart;
                    lengths[count] = nameLength;
                    count++;

                    while (position < tag.End && IsAsciiWhitespace(_buffer[position])) {
                        position++;
                    }
                    if (position >= tag.End || _buffer[position++] != (byte)'=') {
                        return false;
                    }
                    while (position < tag.End && IsAsciiWhitespace(_buffer[position])) {
                        position++;
                    }
                    if (position >= tag.End || _buffer[position] is not ((byte)'\"') and not ((byte)'\'')) {
                        return false;
                    }
                    byte quote = _buffer[position++];
                    int valueStart = position;
                    while (position < tag.End && _buffer[position] != quote) {
                        if (_buffer[position] == (byte)'<') {
                            return false;
                        }
                        position++;
                    }
                    if (position >= tag.End) {
                        return false;
                    }
                    int valueLength = position - valueStart;
                    position++;

                    if (AsciiEquals(nameStart, nameLength, "xmlns")) {
                        declaresDefaultNamespace = true;
                        declaresSpreadsheetNamespace =
                            AsciiEquals(valueStart, valueLength, SpreadsheetNamespace)
                            || AsciiEquals(valueStart, valueLength, StrictSpreadsheetNamespace);
                    }
                }
                return true;
            }

            private bool ValidateCanonicalNamespaceUsage(
                Utf8Tag tag,
                bool isRootStartTag,
                Span<int> namespacePrefixStarts,
                Span<int> namespacePrefixLengths,
                ref int namespacePrefixCount) {
                Span<int> usedPrefixStarts = stackalloc int[MaximumFastXmlAttributes];
                Span<int> usedPrefixLengths = stackalloc int[MaximumFastXmlAttributes];
                int usedPrefixCount = 0;
                int position = tag.NameEnd;
                while (position < tag.End) {
                    while (position < tag.End && IsAsciiWhitespace(_buffer![position])) {
                        position++;
                    }
                    if (position >= tag.End || _buffer![position] == (byte)'/') {
                        break;
                    }

                    int nameStart = position;
                    int colon = -1;
                    while (position < tag.End && !IsAttributeNameTerminator(_buffer[position])) {
                        if (_buffer[position] == (byte)':') {
                            if (colon >= 0) {
                                return false;
                            }
                            colon = position;
                        }
                        position++;
                    }
                    int nameLength = position - nameStart;
                    while (position < tag.End && IsAsciiWhitespace(_buffer[position])) {
                        position++;
                    }
                    if (position >= tag.End || _buffer[position++] != (byte)'=') {
                        return false;
                    }
                    while (position < tag.End && IsAsciiWhitespace(_buffer[position])) {
                        position++;
                    }
                    if (position >= tag.End || _buffer[position] is not ((byte)'\"') and not ((byte)'\'')) {
                        return false;
                    }
                    byte quote = _buffer[position++];
                    int valueStart = position;
                    while (position < tag.End && _buffer[position] != quote) {
                        position++;
                    }
                    if (position >= tag.End) {
                        return false;
                    }
                    int valueLength = position - valueStart;
                    position++;

                    if (colon < 0) {
                        continue;
                    }

                    int prefixLength = colon - nameStart;
                    int localStart = colon + 1;
                    int localLength = nameStart + nameLength - localStart;
                    if (AsciiEquals(nameStart, prefixLength, "xmlns")) {
                        if (!isRootStartTag
                            || localLength == 0
                            || valueLength == 0
                            || AsciiEquals(localStart, localLength, "xml")
                            || AsciiEquals(localStart, localLength, "xmlns")
                            || namespacePrefixCount == namespacePrefixStarts.Length) {
                            return false;
                        }
                        namespacePrefixStarts[namespacePrefixCount] = localStart;
                        namespacePrefixLengths[namespacePrefixCount] = localLength;
                        namespacePrefixCount++;
                        continue;
                    }

                    if (AsciiEquals(nameStart, prefixLength, "xml")) {
                        continue;
                    }
                    if (usedPrefixCount == usedPrefixStarts.Length) {
                        return false;
                    }
                    usedPrefixStarts[usedPrefixCount] = nameStart;
                    usedPrefixLengths[usedPrefixCount] = prefixLength;
                    usedPrefixCount++;
                }

                for (int usedIndex = 0; usedIndex < usedPrefixCount; usedIndex++) {
                    bool bound = false;
                    for (int namespaceIndex = 0; namespaceIndex < namespacePrefixCount; namespaceIndex++) {
                        if (ByteRangesEqual(
                                usedPrefixStarts[usedIndex],
                                usedPrefixLengths[usedIndex],
                                namespacePrefixStarts[namespaceIndex],
                                namespacePrefixLengths[namespaceIndex])) {
                            bound = true;
                            break;
                        }
                    }
                    if (!bound) {
                        return false;
                    }
                }
                return true;
            }

            private bool IsValidUnprefixedXmlName(int start, int end) {
                if (start >= end || !IsAsciiXmlNameStart(_buffer![start])) {
                    return false;
                }
                for (int index = start + 1; index < end; index++) {
                    byte value = _buffer![index];
                    if (value == (byte)':' || !IsAsciiXmlNameCharacter(value)) {
                        return false;
                    }
                }
                return true;
            }

            private bool IsValidXmlAttributeName(int start, int end) {
                if (start >= end || !IsAsciiXmlNameStart(_buffer![start])) {
                    return false;
                }
                for (int index = start + 1; index < end; index++) {
                    byte value = _buffer![index];
                    if (value != (byte)':' && !IsAsciiXmlNameCharacter(value)) {
                        return false;
                    }
                }
                return true;
            }

            private static bool IsAsciiXmlNameStart(byte value) =>
                value is >= (byte)'A' and <= (byte)'Z'
                or >= (byte)'a' and <= (byte)'z'
                or (byte)'_';

            private static bool IsAsciiXmlNameCharacter(byte value) =>
                IsAsciiXmlNameStart(value)
                || value is >= (byte)'0' and <= (byte)'9'
                or (byte)'.'
                or (byte)'-';

            private bool ByteRangesEqual(int firstStart, int firstLength, int secondStart, int secondLength) =>
                firstLength == secondLength
                && _buffer!.AsSpan(firstStart, firstLength).SequenceEqual(
                    _buffer.AsSpan(secondStart, secondLength));

            private bool ContainsOnlyAsciiWhitespace(int start, int end) {
                for (int index = start; index < end; index++) {
                    if (!IsAsciiWhitespace(_buffer![index])) {
                        return false;
                    }
                }
                return true;
            }
        }
    }
}
