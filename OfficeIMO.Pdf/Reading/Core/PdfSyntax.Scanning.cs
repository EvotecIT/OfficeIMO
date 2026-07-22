namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static int FindObjectEnd(
        string text,
        int start,
        IReadOnlyDictionary<(int ObjectNumber, int Generation), int>? declaredLengthValues = null,
        PdfReadLimits? limits = null,
        int? maximumIndex = null) {
        int limit = Math.Min(text.Length, maximumIndex ?? text.Length);
        int searchFrom = start;
        while (searchFrom >= 0 && searchFrom < limit) {
            int streamIdx = IndexOfKeywordOutsideLiteralString(text, "stream", searchFrom, limit);
            int endObjIdx = IndexOfKeywordOutsideLiteralString(text, "endobj", searchFrom, limit);

            if (streamIdx < 0) {
                return endObjIdx < 0 ? -1 : endObjIdx + 6;
            }

            if (endObjIdx >= 0 && endObjIdx < streamIdx) {
                return endObjIdx + 6;
            }

            int afterStream = SkipEOL(text, streamIdx + 6, limit);
            if (TryFindDeclaredStreamObjectEnd(
                    text,
                    start,
                    streamIdx,
                    afterStream,
                    declaredLengthValues,
                    limits,
                    limit,
                    out int declaredObjectEnd)) {
                return declaredObjectEnd;
            }

            int endStreamSearchFrom = afterStream;
            while (endStreamSearchFrom < limit) {
                int endStreamIdx = IndexOfKeyword(text, "endstream", endStreamSearchFrom, limit);
                if (endStreamIdx < 0) {
                    return -1;
                }

                // Without a trustworthy inline /Length, a later object header is the
                // safest recovery boundary. Otherwise a malformed stream can consume a
                // complete following object through that object's endstream/endobj pair.
                if (ContainsIndirectObjectHeader(text, endStreamSearchFrom, endStreamIdx)) {
                    return -1;
                }

                int nextToken = SkipWhitespaceAndComments(text, endStreamIdx + 9, limit);
                if (IsKeywordAt(text, "endobj", nextToken, limit)) {
                    return nextToken + 6;
                }

                if (IsIndirectObjectHeaderAt(text, nextToken)) {
                    return -1;
                }

                // Binary stream data may contain the bytes "endstream". Only a
                // structurally bounded delimiter is accepted.
                endStreamSearchFrom = endStreamIdx + 9;
            }
        }

        return -1;
    }

    private static int SkipWhitespaceAndComments(string text, int index, int limit) {
        while (index < limit) {
            while (index < limit && char.IsWhiteSpace(text[index])) {
                index++;
            }

            if (index >= limit || text[index] != '%') {
                return index;
            }

            while (index < limit && text[index] != '\r' && text[index] != '\n') {
                index++;
            }
        }

        return index;
    }

    private static bool IsKeywordAt(string text, string keyword, int index, int limit) =>
        index >= 0 &&
        index + keyword.Length <= limit &&
        string.CompareOrdinal(text, index, keyword, 0, keyword.Length) == 0 &&
        HasKeywordBoundary(text, index - 1, 0, limit) &&
        HasKeywordBoundary(text, index + keyword.Length, 0, limit);

    private static bool TryFindDeclaredStreamObjectEnd(
        string text,
        int objectStart,
        int streamIndex,
        int dataStart,
        IReadOnlyDictionary<(int ObjectNumber, int Generation), int>? declaredLengthValues,
        PdfReadLimits? limits,
        int limit,
        out int objectEnd) {
        objectEnd = -1;
        int dictionaryStart = text.IndexOf("<<", objectStart, streamIndex - objectStart, StringComparison.Ordinal);
        if (dictionaryStart < 0) {
            return false;
        }

        int dictionaryEnd = FindDictEnd(text, dictionaryStart, streamIndex);
        if (dictionaryEnd < 0) {
            return false;
        }

        int dictionaryCharacters = dictionaryEnd - dictionaryStart - 2;
        int maximumDictionaryCharacters = limits?.MaxObjectCharacters ?? 1_000_000;
        if (dictionaryCharacters > maximumDictionaryCharacters) {
            if (limits != null) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.ObjectCharacters,
                    maximumDictionaryCharacters,
                    dictionaryCharacters);
            }
            return false;
        }

        PdfDictionary? dictionary;
        try {
            string dictionaryText = text.Substring(dictionaryStart + 2, dictionaryCharacters);
            dictionary = limits == null
                ? ParseDictionary(dictionaryText)
                : ParseDictionary(dictionaryText, limits);
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            return false;
        }

        if (!TryResolveDeclaredStreamLength(dictionary, declaredLengthValues, out int byteLength)) {
            return false;
        }

        if (!TryGetDeclaredEndStreamIndex(text, dataStart, byteLength, limit, out int endStream)) {
            return false;
        }

        int endObject = SkipWhitespaceAndComments(text, endStream + 9, limit);
        if (!IsKeywordAt(text, "endobj", endObject, limit)) {
            return false;
        }

        objectEnd = endObject + 6;
        return true;
    }

    private static bool TryResolveDeclaredStreamLength(
        PdfDictionary dictionary,
        IReadOnlyDictionary<(int ObjectNumber, int Generation), int>? declaredLengthValues,
        out int byteLength) {
        byteLength = -1;
        if (dictionary.Get<PdfNumber>("Length") is PdfNumber directLength) {
            return TryNormalizeStreamLength(directLength.Value, out byteLength);
        }

        if (dictionary.Get<PdfReference>("Length") is not PdfReference lengthReference ||
            declaredLengthValues == null ||
            !declaredLengthValues.TryGetValue(
                (lengthReference.ObjectNumber, lengthReference.Generation),
                out byteLength)) {
            return false;
        }

        return byteLength >= 0;
    }

    private static Dictionary<(int ObjectNumber, int Generation), int> BuildDeclaredLengthValueIndex(
        string text,
        IReadOnlyList<IndirectObjectHeader> objectMatches,
        System.Diagnostics.Stopwatch parseTimer,
        PdfReadLimits limits) {
        var streamRanges = new List<(int Start, int End)>();
        var knownStreamRanges = new HashSet<(int Start, int End)>();
        DiscoverDeclaredStreamRanges(
            text,
            objectMatches,
            streamRanges,
            knownStreamRanges,
            declaredLengthValues: null,
            parseTimer,
            limits);
        Dictionary<(int ObjectNumber, int Generation), int> values =
            BuildScalarObjectIndex(text, objectMatches, streamRanges, parseTimer, limits);

        // A later indirect-length stream can hide scalar-looking bytes that initially
        // appear to be objects. Iterate to a timed fixed point: every successful pass
        // expands the covered stream ranges, whose identities are bounded by the
        // already-enforced indirect-object budget.
        while (true) {
            ThrowIfParsingTimeExceeded(parseTimer, limits);
            bool coverageChanged = DiscoverDeclaredStreamRanges(
                text,
                objectMatches,
                streamRanges,
                knownStreamRanges,
                values,
                parseTimer,
                limits);
            if (!coverageChanged) {
                return values;
            }

            values = BuildScalarObjectIndex(
                text,
                objectMatches,
                streamRanges,
                parseTimer,
                limits);
        }
    }

    private static bool DiscoverDeclaredStreamRanges(
        string text,
        IReadOnlyList<IndirectObjectHeader> objectMatches,
        List<(int Start, int End)> streamRanges,
        HashSet<(int Start, int End)> knownStreamRanges,
        IReadOnlyDictionary<(int ObjectNumber, int Generation), int>? declaredLengthValues,
        System.Diagnostics.Stopwatch parseTimer,
        PdfReadLimits limits) {
        var discoveredRanges = new List<(int Start, int End)>();
        for (int i = 0; i < objectMatches.Count; i++) {
            if ((i & 127) == 0) {
                ThrowIfParsingTimeExceeded(parseTimer, limits);
            }

            var match = objectMatches[i];
            if (IsInsideStreamRange(match.Index, streamRanges) ||
                IsInsideStreamRange(match.Index, discoveredRanges) ||
                !TryReadDeclaredStreamRange(
                    text,
                    match.Index + match.Length,
                    declaredLengthValues,
                    limits,
                    out (int Start, int End) streamRange) ||
                !knownStreamRanges.Add(streamRange)) {
                continue;
            }

            AddOrderedStreamRange(discoveredRanges, streamRange);
        }

        return MergeStreamRanges(streamRanges, discoveredRanges);
    }

    private static bool TryReadDeclaredStreamRange(
        string text,
        int bodyStart,
        IReadOnlyDictionary<(int ObjectNumber, int Generation), int>? declaredLengthValues,
        PdfReadLimits limits,
        out (int Start, int End) streamRange) {
        streamRange = default;
        int dictionaryStart = SkipWhitespaceAndComments(text, bodyStart, text.Length);
        if (dictionaryStart + 1 >= text.Length ||
            text[dictionaryStart] != '<' ||
            text[dictionaryStart + 1] != '<') {
            return false;
        }

        int dictionaryEnd = FindDictEnd(text, dictionaryStart, text.Length);
        if (dictionaryEnd < 0 ||
            dictionaryEnd - dictionaryStart - 2 > limits.MaxObjectCharacters) {
            return false;
        }

        PdfDictionary? dictionary;
        try {
            dictionary = ParseDictionary(text.Substring(
                dictionaryStart + 2,
                dictionaryEnd - dictionaryStart - 2), limits);
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            return false;
        }

        int streamIndex = SkipWhitespaceAndComments(text, dictionaryEnd, text.Length);
        if (!IsKeywordAt(text, "stream", streamIndex, text.Length) ||
            !TryResolveDeclaredStreamLength(dictionary, declaredLengthValues, out int byteLength)) {
            return false;
        }

        int dataStart = SkipEOL(text, streamIndex + 6, text.Length);
        if (!TryGetDeclaredEndStreamIndex(
                text,
                dataStart,
                byteLength,
                text.Length,
                out int endStream)) {
            return false;
        }

        int endObject = SkipWhitespaceAndComments(text, endStream + 9, text.Length);
        if (!IsKeywordAt(text, "endobj", endObject, text.Length)) {
            return false;
        }

        streamRange = (dataStart, endStream);
        return true;
    }

    private static Dictionary<(int ObjectNumber, int Generation), int> BuildScalarObjectIndex(
        string text,
        IReadOnlyList<IndirectObjectHeader> objectMatches,
        List<(int Start, int End)> streamRanges,
        System.Diagnostics.Stopwatch parseTimer,
        PdfReadLimits limits) {
        var values = new Dictionary<(int ObjectNumber, int Generation), int>();
        for (int i = 0; i < objectMatches.Count; i++) {
            if ((i & 127) == 0) {
                ThrowIfParsingTimeExceeded(parseTimer, limits);
            }

            IndirectObjectHeader match = objectMatches[i];
            if (IsInsideStreamRange(match.Index, streamRanges) ||
                match.ObjectNumber < 0 ||
                match.Generation < 0) {
                continue;
            }

            int bodyStart = match.Index + match.Length;
            int objectLimit = i + 1 < objectMatches.Count
                ? objectMatches[i + 1].Index
                : text.Length;
            int endObject = IndexOfKeywordOutsideLiteralString(text, "endobj", bodyStart, objectLimit);
            int bodyLength = endObject - bodyStart;
            if (endObject < 0 || bodyLength < 0 || bodyLength > 256) {
                continue;
            }

            List<PdfToken> tokens;
            try {
                tokens = Tokenize(text.Substring(bodyStart, bodyLength));
            } catch (Exception exception) when (exception is not OutOfMemoryException) {
                continue;
            }

            if (tokens.Count == 1 &&
                double.TryParse(
                    tokens[0].Text,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double value) &&
                TryNormalizeStreamLength(value, out int byteLength)) {
                values[(match.ObjectNumber, match.Generation)] = byteLength;
            }
        }

        return values;
    }

    private static void AddOrderedStreamRange(
        List<(int Start, int End)> ranges,
        (int Start, int End) candidate) {
        if (ranges.Count == 0 || candidate.Start > ranges[ranges.Count - 1].End) {
            ranges.Add(candidate);
            return;
        }

        (int Start, int End) last = ranges[ranges.Count - 1];
        if (candidate.End > last.End) {
            ranges[ranges.Count - 1] = (last.Start, candidate.End);
        }
    }

    private static bool MergeStreamRanges(
        List<(int Start, int End)> ranges,
        List<(int Start, int End)> additions) {
        if (additions.Count == 0) {
            return false;
        }

        var merged = new List<(int Start, int End)>(ranges.Count + additions.Count);
        int existingIndex = 0;
        int additionIndex = 0;
        while (existingIndex < ranges.Count || additionIndex < additions.Count) {
            (int Start, int End) candidate;
            if (additionIndex >= additions.Count ||
                (existingIndex < ranges.Count &&
                 ranges[existingIndex].Start <= additions[additionIndex].Start)) {
                candidate = ranges[existingIndex++];
            } else {
                candidate = additions[additionIndex++];
            }

            AddOrderedStreamRange(merged, candidate);
        }

        bool changed = merged.Count != ranges.Count;
        if (!changed) {
            for (int i = 0; i < merged.Count; i++) {
                if (merged[i] != ranges[i]) {
                    changed = true;
                    break;
                }
            }
        }

        if (changed) {
            ranges.Clear();
            ranges.AddRange(merged);
        }

        return changed;
    }

    private static bool IsInsideStreamRange(int offset, List<(int Start, int End)> streamRanges) {
        int low = 0;
        int high = streamRanges.Count - 1;
        while (low <= high) {
            int middle = low + ((high - low) / 2);
            (int Start, int End) range = streamRanges[middle];
            if (offset < range.Start) {
                high = middle - 1;
            } else if (offset >= range.End) {
                low = middle + 1;
            } else {
                return true;
            }
        }

        return false;
    }

    private static bool TryNormalizeStreamLength(double value, out int byteLength) {
        byteLength = -1;
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            return false;
        }

        byteLength = (int)Math.Max(0D, Math.Min(int.MaxValue, value));
        return true;
    }

    private static bool TryGetDeclaredEndStreamIndex(
        string text,
        int dataStart,
        int byteLength,
        int limit,
        out int endStream) {
        endStream = -1;
        if (byteLength < 0 || dataStart < 0 || dataStart > limit - byteLength) {
            return false;
        }

        int candidate = dataStart + byteLength;
        if (candidate < limit && text[candidate] == '\r') candidate++;
        if (candidate < limit && text[candidate] == '\n') candidate++;
        if (!IsKeywordAt(text, "endstream", candidate, limit)) {
            return false;
        }

        endStream = candidate;
        return true;
    }

    private static List<IndirectObjectHeader> FindIndirectObjectHeaders(
        string text,
        System.Diagnostics.Stopwatch parseTimer,
        PdfReadLimits limits) {
        var headers = new List<IndirectObjectHeader>();
        int cursor = 0;
        while (TryFindIndirectObjectHeader(
            text,
            cursor,
            text.Length,
            out IndirectObjectHeader header,
            parseTimer,
            limits)) {
            headers.Add(header);
            if (headers.Count > limits.MaxIndirectObjects) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.IndirectObjects,
                    limits.MaxIndirectObjects,
                    headers.Count);
            }

            cursor = header.Index + header.Length;
        }

        ThrowIfParsingTimeExceeded(parseTimer, limits);
        return headers;
    }

    private static bool TryFindIndirectObjectHeader(
        string text,
        int start,
        int limit,
        out IndirectObjectHeader header,
        System.Diagnostics.Stopwatch? parseTimer = null,
        PdfReadLimits? limits = null) {
        header = default;
        if (start < 0) start = 0;
        if (limit > text.Length) limit = text.Length;
        if (start >= limit) return false;

        int index = start;
        while (index < limit) {
            if ((index & 0x3FFF) == 0 && parseTimer is not null && limits is not null) {
                ThrowIfParsingTimeExceeded(parseTimer, limits);
            }

            if (!char.IsDigit(text[index])) {
                index++;
                continue;
            }

            if (TryReadIndirectObjectHeaderAt(
                text,
                index,
                limit,
                out header,
                parseTimer,
                limits)) {
                return true;
            }

            // A failed header can still begin with a very large digit run. Skip the
            // whole run so malformed or signature-reservation data stays linear.
            do {
                index++;
            } while (index < limit && char.IsDigit(text[index]));
        }

        return false;
    }

    private static bool TryReadIndirectObjectHeaderAt(
        string text,
        int index,
        int limit,
        out IndirectObjectHeader header,
        System.Diagnostics.Stopwatch? parseTimer = null,
        PdfReadLimits? limits = null) {
        header = default;
        if (index < 0 ||
            index >= limit ||
            limit > text.Length ||
            !char.IsDigit(text[index]) ||
            (index > 0 && char.IsDigit(text[index - 1]))) {
            return false;
        }

        if (!TryReadNonNegativeInteger(
                text,
                index,
                limit,
                out int objectNumber,
                out int cursor,
                parseTimer,
                limits) ||
            cursor >= limit ||
            !char.IsWhiteSpace(text[cursor])) {
            return false;
        }

        SkipIndirectObjectHeaderWhitespace(
            text,
            ref cursor,
            limit,
            parseTimer,
            limits);

        if (!TryReadNonNegativeInteger(
                text,
                cursor,
                limit,
                out int generation,
                out cursor,
                parseTimer,
                limits) ||
            cursor >= limit ||
            !char.IsWhiteSpace(text[cursor])) {
            return false;
        }

        SkipIndirectObjectHeaderWhitespace(
            text,
            ref cursor,
            limit,
            parseTimer,
            limits);

        if (cursor > limit - 3 ||
            text[cursor] != 'o' ||
            text[cursor + 1] != 'b' ||
            text[cursor + 2] != 'j') {
            return false;
        }

        int end = cursor + 3;
        header = new IndirectObjectHeader(
            index,
            end - index,
            objectNumber,
            generation);
        return true;
    }

    private static bool TryReadNonNegativeInteger(
        string text,
        int index,
        int limit,
        out int value,
        out int end,
        System.Diagnostics.Stopwatch? parseTimer = null,
        PdfReadLimits? limits = null) {
        value = 0;
        end = index;
        if (index < 0 || index >= limit || !char.IsDigit(text[index])) {
            return false;
        }

        int parsed = 0;
        bool overflow = false;
        while (end < limit && char.IsDigit(text[end])) {
            if ((end & 0x3FFF) == 0 && parseTimer is not null && limits is not null) {
                ThrowIfParsingTimeExceeded(parseTimer, limits);
            }

            if (!overflow) {
                int digit = text[end] - '0';
                if (parsed > (int.MaxValue - digit) / 10) {
                    overflow = true;
                } else {
                    parsed = parsed * 10 + digit;
                }
            }
            end++;
        }

        if (overflow) {
            return false;
        }

        value = parsed;
        return true;
    }

    private static void SkipIndirectObjectHeaderWhitespace(
        string text,
        ref int index,
        int limit,
        System.Diagnostics.Stopwatch? parseTimer,
        PdfReadLimits? limits) {
        while (index < limit && char.IsWhiteSpace(text[index])) {
            if ((index & 0x3FFF) == 0 && parseTimer is not null && limits is not null) {
                ThrowIfParsingTimeExceeded(parseTimer, limits);
            }

            index++;
        }
    }

    private readonly struct IndirectObjectHeader {
        internal IndirectObjectHeader(
            int index,
            int length,
            int objectNumber,
            int generation) {
            Index = index;
            Length = length;
            ObjectNumber = objectNumber;
            Generation = generation;
        }

        internal int Index { get; }
        internal int Length { get; }
        internal int ObjectNumber { get; }
        internal int Generation { get; }
    }

    private static bool ContainsIndirectObjectHeader(string text, int start, int limit) {
        return TryFindIndirectObjectHeader(text, start, limit, out _);
    }

    private static bool IsIndirectObjectHeaderAt(string text, int index) {
        return TryReadIndirectObjectHeaderAt(text, index, text.Length, out _);
    }

    private static int IndexOfKeywordOutsideLiteralString(string text, string keyword, int start, int limit) {
        if (start < 0) start = 0;
        if (limit > text.Length) limit = text.Length;

        int literalDepth = 0;
        bool escaped = false;
        for (int i = start; i < limit; i++) {
            char c = text[i];
            if (literalDepth > 0) {
                if (escaped) {
                    escaped = false;
                    continue;
                }

                if (c == '\\') {
                    escaped = true;
                    continue;
                }

                if (c == '(') {
                    literalDepth++;
                    continue;
                }

                if (c == ')') {
                    literalDepth--;
                }

                continue;
            }

            if (c == '(') {
                literalDepth = 1;
                continue;
            }

            if (i + keyword.Length <= limit &&
                string.CompareOrdinal(text, i, keyword, 0, keyword.Length) == 0 &&
                HasKeywordBoundary(text, i - 1, start, limit) &&
                HasKeywordBoundary(text, i + keyword.Length, start, limit)) {
                return i;
            }
        }

        return -1;
    }

    private static int FindDictEnd(string text, int dictStart, int limit) {
        int depth = 0;
        for (int i = dictStart; i + 1 < limit; i++) {
            char c = text[i]; char n = text[i + 1];
            if (c == '<' && n == '<') { depth++; i++; continue; }
            if (c == '>' && n == '>') { depth--; i++; if (depth == 0) return i + 1; continue; }
        }
        return -1;
    }

    private static int IndexOfKeyword(string text, string keyword, int start, int limit) {
        if (start < 0) start = 0;
        if (limit > text.Length) limit = text.Length;

        int searchFrom = start;
        while (searchFrom < limit) {
            int idx = text.IndexOf(keyword, searchFrom, StringComparison.Ordinal);
            if (idx < 0 || idx >= limit) {
                return -1;
            }

            int end = idx + keyword.Length;
            if (HasKeywordBoundary(text, idx - 1, start, limit) &&
                HasKeywordBoundary(text, end, start, limit)) {
                return idx;
            }

            searchFrom = idx + 1;
        }

        return -1;
    }

    private static bool HasKeywordBoundary(string text, int idx, int start, int limit) {
        if (idx < start || idx >= limit) {
            return true;
        }

        char c = text[idx];
        return char.IsWhiteSpace(c) || c is '<' or '>' or '[' or ']' or '%';
    }

    private static int SkipEOL(string text, int idx, int limit) {
        if (idx < limit) {
            if (text[idx] == '\r') idx++;
            if (idx < limit && text[idx] == '\n') idx++;
        }
        return idx;
    }

    private static string SafeSlice(string s, int start, int length, int maxLen) {
        int len = Math.Min(length, Math.Max(0, maxLen));
        if (start < 0) start = 0;
        if (start + len > s.Length) len = s.Length - start;
        if (len <= 0) return string.Empty;
        return s.Substring(start, len);
    }
}
