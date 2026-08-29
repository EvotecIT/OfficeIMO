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
        if (StartsWith(bytes, xref, "xref")) return true;
        return IsXrefStreamObject(bytes, xref, startXref);
    }

    private static bool IsXrefStreamObject(byte[] bytes, int offset, int limit) {
        if (!TryReadDecimal(bytes, ref offset, limit, out _)) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!TryReadDecimal(bytes, ref offset, limit, out _)) return false;
        SkipWhitespace(bytes, ref offset, limit);
        if (!StartsWith(bytes, offset, "obj")) return false;
        int dictionaryEnd = IndexOf(bytes, "stream", offset + 3, Math.Min(limit, offset + 4096));
        return dictionaryEnd >= 0
            && IndexOf(bytes, "/Type", offset + 3, dictionaryEnd) >= 0
            && IndexOf(bytes, "/XRef", offset + 3, dictionaryEnd) >= 0;
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
