namespace OfficeIMO.IWork.Internal;

internal static class IWorkPdfInfo {
    internal static bool IsComplete(byte[] bytes) {
        if (bytes.Length < 20 || !StartsWith(bytes, 0, "%PDF-")) return false;

        int eof = LastIndexOf(bytes, "%%EOF");
        if (eof < 0 || !ContainsOnlyTrailingWhitespace(bytes, eof + 5)) return false;
        int startXref = LastIndexOf(bytes, "startxref", eof);
        if (startXref < 0) return false;

        int offset = startXref + 9;
        SkipWhitespace(bytes, ref offset, eof);
        if (!TryReadDecimal(bytes, ref offset, eof, out long xrefOffset)
            || xrefOffset < 0 || xrefOffset >= startXref || xrefOffset > int.MaxValue) return false;

        int xref = (int)xrefOffset;
        SkipWhitespace(bytes, ref xref, startXref);
        if (StartsWith(bytes, xref, "xref")) return IsClassicXref(bytes, xref, startXref);
        return IsXrefStreamObject(bytes, xref, startXref);
    }

    private static bool IsClassicXref(byte[] bytes, int offset, int limit) {
        offset += 4;
        var inUseOffsets = new Dictionary<(long Object, long Generation), int>();
        bool hasSubsection = false;
        while (offset < limit) {
            SkipWhitespace(bytes, ref offset, limit);
            if (StartsWith(bytes, offset, "trailer")) break;
            if (!TryReadDecimal(bytes, ref offset, limit, out long firstObject)) return false;
            SkipWhitespace(bytes, ref offset, limit);
            if (!TryReadDecimal(bytes, ref offset, limit, out long entryCount)
                || entryCount <= 0 || entryCount > 1_000_000
                || firstObject > long.MaxValue - (entryCount - 1)) return false;
            hasSubsection = true;
            for (long index = 0; index < entryCount; index++) {
                SkipWhitespace(bytes, ref offset, limit);
                if (!TryReadFixedDecimal(bytes, ref offset, limit, 10, out long objectOffset)) return false;
                SkipHorizontalWhitespace(bytes, ref offset, limit);
                if (!TryReadFixedDecimal(bytes, ref offset, limit, 5, out long generation)) return false;
                SkipHorizontalWhitespace(bytes, ref offset, limit);
                if (offset >= limit || bytes[offset] != (byte)'n' && bytes[offset] != (byte)'f') return false;
                bool inUse = bytes[offset++] == (byte)'n';
                if (!ConsumeLineEnd(bytes, ref offset, limit)) return false;
                if (inUse && objectOffset <= int.MaxValue) {
                    inUseOffsets[(checked(firstObject + index), generation)] = (int)objectOffset;
                }
            }
        }
        if (!hasSubsection || !StartsWith(bytes, offset, "trailer")) return false;
        int dictionaryStart = IndexOf(bytes, "<<", offset + 7, Math.Min(limit, offset + 65536));
        if (dictionaryStart < 0) return false;
        int dictionaryEnd = IndexOf(bytes, ">>", dictionaryStart + 2, Math.Min(limit, dictionaryStart + 65536));
        if (dictionaryEnd < 0
            || !TryReadDictionaryInteger(bytes, dictionaryStart, dictionaryEnd, "/Size", out long size)
            || size <= 0
            || !TryReadDictionaryReference(bytes, dictionaryStart, dictionaryEnd, "/Root",
                out long rootObject, out long rootGeneration)
            || rootObject <= 0 || rootObject >= size
            || !inUseOffsets.TryGetValue((rootObject, rootGeneration), out int rootOffset)) return false;
        if (!IsCatalogObjectAt(bytes, rootOffset, offset, rootObject, rootGeneration,
                out long pagesObject, out long pagesGeneration)
            || !inUseOffsets.TryGetValue((pagesObject, pagesGeneration), out int pagesOffset)) return false;
        return IsPagesObjectAt(bytes, pagesOffset, offset, pagesObject, pagesGeneration);
    }

    private static bool IsXrefStreamObject(byte[] bytes, int offset, int limit) {
        if (!TryReadDecimal(bytes, ref offset, limit, out _)) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!TryReadDecimal(bytes, ref offset, limit, out _)) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!StartsWith(bytes, offset, "obj")) return false;
        int dictionaryStart = IndexOf(bytes, "<<", offset + 3, Math.Min(limit, offset + 4096));
        int dictionaryEnd = dictionaryStart < 0
            ? -1
            : IndexOf(bytes, ">>", dictionaryStart + 2, Math.Min(limit, dictionaryStart + 65536));
        int streamStart = dictionaryEnd < 0 ? -1 : dictionaryEnd + 2;
        if (streamStart >= 0) SkipWhitespace(bytes, ref streamStart, limit);
        if (dictionaryEnd < 0
            || streamStart < 0 || !StartsWith(bytes, streamStart, "stream")
            || IndexOf(bytes, "endstream", streamStart + 6, limit) < 0
            || !HasDictionaryNameValue(bytes, dictionaryStart, dictionaryEnd, "/Type", "/XRef")
            || !HasXrefWidths(bytes, dictionaryStart, dictionaryEnd)
            || !TryReadDictionaryInteger(bytes, dictionaryStart, dictionaryEnd, "/Size", out long size)
            || size <= 0
            || !TryReadDictionaryReference(bytes, dictionaryStart, dictionaryEnd, "/Root",
                out long rootObject, out long rootGeneration)
            || rootObject <= 0 || rootObject >= size) return false;
        int rootOffset = FindObjectHeader(bytes, rootObject, rootGeneration, offset);
        if (rootOffset < 0 || !IsCatalogObjectAt(bytes, rootOffset, offset, rootObject, rootGeneration,
                out long pagesObject, out long pagesGeneration)) return false;
        int pagesOffset = FindObjectHeader(bytes, pagesObject, pagesGeneration, offset);
        return pagesOffset >= 0 && IsPagesObjectAt(bytes, pagesOffset, offset, pagesObject, pagesGeneration);
    }

    private static bool TryReadFixedDecimal(byte[] bytes, ref int offset, int limit, int digits, out long value) {
        value = 0;
        if (offset < 0 || offset > limit - digits) return false;
        for (int index = 0; index < digits; index++) {
            byte current = bytes[offset++];
            if (current < (byte)'0' || current > (byte)'9') return false;
            value = value * 10 + current - (byte)'0';
        }
        return true;
    }

    private static bool TryReadDictionaryInteger(byte[] bytes, int start, int end,
        string name, out long value) {
        value = 0;
        int offset = IndexOfDictionaryName(bytes, name, start, end);
        if (offset < 0) return false;
        offset += name.Length;
        SkipWhitespace(bytes, ref offset, end);
        if (!TryReadDecimal(bytes, ref offset, end, out value)) return false;
        SkipWhitespace(bytes, ref offset, end);
        return offset >= end || IsDelimiter(bytes[offset]);
    }

    private static bool TryReadDictionaryReference(byte[] bytes, int start, int end,
        string name, out long objectNumber, out long generation) {
        objectNumber = 0;
        generation = 0;
        int offset = IndexOfDictionaryName(bytes, name, start, end);
        if (offset < 0) return false;
        offset += name.Length;
        SkipWhitespace(bytes, ref offset, end);
        if (!TryReadDecimal(bytes, ref offset, end, out objectNumber)) return false;
        SkipWhitespace(bytes, ref offset, end);
        if (!TryReadDecimal(bytes, ref offset, end, out generation)) return false;
        SkipWhitespace(bytes, ref offset, end);
        if (offset >= end || bytes[offset++] != (byte)'R') return false;
        return offset >= end || IsDelimiter(bytes[offset]);
    }

    private static bool IsCatalogObjectAt(byte[] bytes, int offset, int limit,
        long expectedObject, long expectedGeneration, out long pagesObject, out long pagesGeneration) {
        pagesObject = 0;
        pagesGeneration = 0;
        if (!TryGetObjectDictionary(bytes, ref offset, limit, expectedObject, expectedGeneration,
                out int dictionaryStart, out int dictionaryEnd)) return false;
        return HasDictionaryNameValue(bytes, dictionaryStart, dictionaryEnd, "/Type", "/Catalog")
            && TryReadDictionaryReference(bytes, dictionaryStart, dictionaryEnd, "/Pages",
                out pagesObject, out pagesGeneration);
    }

    private static bool IsPagesObjectAt(byte[] bytes, int offset, int limit,
        long expectedObject, long expectedGeneration) {
        if (!TryGetObjectDictionary(bytes, ref offset, limit, expectedObject, expectedGeneration,
                out int dictionaryStart, out int dictionaryEnd)) return false;
        return HasDictionaryNameValue(bytes, dictionaryStart, dictionaryEnd, "/Type", "/Pages")
            && IndexOfDictionaryName(bytes, "/Kids", dictionaryStart, dictionaryEnd) >= 0
            && TryReadDictionaryInteger(bytes, dictionaryStart, dictionaryEnd, "/Count", out long count)
            && count >= 0;
    }

    private static bool TryGetObjectDictionary(byte[] bytes, ref int offset, int limit,
        long expectedObject, long expectedGeneration, out int dictionaryStart, out int dictionaryEnd) {
        dictionaryStart = -1;
        dictionaryEnd = -1;
        SkipWhitespace(bytes, ref offset, limit);
        if (!TryReadDecimal(bytes, ref offset, limit, out long objectNumber)
            || objectNumber != expectedObject) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!TryReadDecimal(bytes, ref offset, limit, out long generation)
            || generation != expectedGeneration) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!StartsWith(bytes, offset, "obj")) return false;
        int objectEnd = IndexOf(bytes, "endobj", offset + 3, Math.Min(limit, offset + 65536));
        if (objectEnd < 0) return false;
        dictionaryStart = IndexOf(bytes, "<<", offset + 3, objectEnd);
        dictionaryEnd = dictionaryStart < 0 ? -1 : IndexOf(bytes, ">>", dictionaryStart + 2, objectEnd);
        return dictionaryStart >= 0 && dictionaryEnd >= 0;
    }

    private static bool HasDictionaryNameValue(byte[] bytes, int start, int end,
        string name, string value) {
        int offset = IndexOfDictionaryName(bytes, name, start, end);
        if (offset < 0) return false;
        offset += name.Length;
        SkipWhitespace(bytes, ref offset, end);
        return StartsWith(bytes, offset, value)
            && (offset + value.Length >= end || IsDelimiter(bytes[offset + value.Length]));
    }

    private static bool HasXrefWidths(byte[] bytes, int start, int end) {
        int offset = IndexOfDictionaryName(bytes, "/W", start, end);
        if (offset < 0) return false;
        offset += 2;
        SkipWhitespace(bytes, ref offset, end);
        if (offset >= end || bytes[offset++] != (byte)'[') return false;
        bool hasNonZeroWidth = false;
        for (int index = 0; index < 3; index++) {
            SkipWhitespace(bytes, ref offset, end);
            if (!TryReadDecimal(bytes, ref offset, end, out long width) || width > 8) return false;
            hasNonZeroWidth |= width > 0;
        }
        SkipWhitespace(bytes, ref offset, end);
        return hasNonZeroWidth && offset < end && bytes[offset] == (byte)']';
    }

    private static int IndexOfDictionaryName(byte[] bytes, string name, int start, int end) {
        int offset = Math.Max(0, start);
        while (offset < end) {
            int found = IndexOf(bytes, name, offset, end);
            if (found < 0) return -1;
            int after = found + name.Length;
            if (after >= end || IsDelimiter(bytes[after])) return found;
            offset = found + 1;
        }
        return -1;
    }

    private static bool IsDelimiter(byte value) => IsWhitespace(value)
        || value is (byte)'(' or (byte)')' or (byte)'<' or (byte)'>' or (byte)'[' or (byte)']'
            or (byte)'{' or (byte)'}' or (byte)'/' or (byte)'%';

    private static int FindObjectHeader(byte[] bytes, long objectNumber, long generation, int limit) {
        string header = objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " "
            + generation.ToString(System.Globalization.CultureInfo.InvariantCulture) + " obj";
        int offset = 0;
        while (offset < limit) {
            int found = IndexOf(bytes, header, offset, limit);
            if (found < 0) return -1;
            int after = found + header.Length;
            bool startsAtToken = found == 0 || IsDelimiter(bytes[found - 1]);
            bool endsAtToken = after >= limit || IsDelimiter(bytes[after]);
            if (startsAtToken && endsAtToken) return found;
            offset = found + 1;
        }
        return -1;
    }

    private static void SkipHorizontalWhitespace(byte[] bytes, ref int offset, int limit) {
        while (offset < limit && bytes[offset] is 0x09 or 0x20) offset++;
    }

    private static bool ConsumeLineEnd(byte[] bytes, ref int offset, int limit) {
        SkipHorizontalWhitespace(bytes, ref offset, limit);
        if (offset >= limit) return false;
        if (bytes[offset] == 0x0d) {
            offset++;
            if (offset < limit && bytes[offset] == 0x0a) offset++;
            return true;
        }
        if (bytes[offset] == 0x0a) {
            offset++;
            return true;
        }
        return false;
    }

    private static bool TryReadDecimal(byte[] bytes, ref int offset, int limit, out long value) {
        value = 0;
        int start = offset;
        while (offset < limit && bytes[offset] >= (byte)'0' && bytes[offset] <= (byte)'9') {
            int digit = bytes[offset++] - (byte)'0';
            if (value > (long.MaxValue - digit) / 10) return false;
            value = value * 10 + digit;
        }
        return offset > start;
    }

    private static void SkipWhitespace(byte[] bytes, ref int offset, int limit) {
        while (offset < limit && IsWhitespace(bytes[offset])) offset++;
    }

    private static bool ContainsOnlyTrailingWhitespace(byte[] bytes, int offset) {
        for (int index = offset; index < bytes.Length; index++) {
            if (!IsWhitespace(bytes[index]) && bytes[index] != 0) return false;
        }
        return true;
    }

    private static bool IsWhitespace(byte value) =>
        value is 0x09 or 0x0a or 0x0c or 0x0d or 0x20;

    private static int LastIndexOf(byte[] bytes, string value, int? before = null) {
        int last = -1;
        int limit = Math.Min(before ?? bytes.Length, bytes.Length);
        for (int index = 0; index <= limit - value.Length; index++) {
            if (StartsWith(bytes, index, value)) last = index;
        }
        return last;
    }

    private static int IndexOf(byte[] bytes, string value, int start, int limit) {
        int end = Math.Min(limit, bytes.Length);
        for (int index = Math.Max(0, start); index <= end - value.Length; index++) {
            if (StartsWith(bytes, index, value)) return index;
        }
        return -1;
    }

    private static bool StartsWith(byte[] bytes, int offset, string value) {
        if (offset < 0 || offset > bytes.Length - value.Length) return false;
        for (int index = 0; index < value.Length; index++) {
            if (bytes[offset + index] != (byte)value[index]) return false;
        }
        return true;
    }
}
