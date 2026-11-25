namespace OfficeIMO.Excel {
    /// <summary>
    /// Reads the value of a boolean cell (<c>t="b"</c>). The value is an xsd:boolean, so <c>1</c> and <c>true</c> are
    /// true and surrounding XML whitespace is ignored. <c>true</c> is matched without regard to case, as the rest of the
    /// package reads boolean attributes. Any other text, including <c>0</c> and <c>false</c>, is false.
    /// </summary>
    internal static class ExcelBooleanCellText {
        internal static bool IsTrue(string? text) => text != null && IsTrue(text.AsSpan());

        internal static bool IsTrue(ReadOnlySpan<char> text) {
            int start = 0;
            int end = text.Length;
            while (start < end && IsXmlWhitespace(text[start])) start++;
            while (end > start && IsXmlWhitespace(text[end - 1])) end--;
            int length = end - start;
            if (length == 1) return text[start] == '1';
            if (length != 4) return false;
            return (text[start] | 0x20) == 't'
                && (text[start + 1] | 0x20) == 'r'
                && (text[start + 2] | 0x20) == 'u'
                && (text[start + 3] | 0x20) == 'e';
        }

        internal static bool IsTrue(ReadOnlySpan<byte> text) {
            int start = 0;
            int end = text.Length;
            while (start < end && IsXmlWhitespace((char)text[start])) start++;
            while (end > start && IsXmlWhitespace((char)text[end - 1])) end--;
            int length = end - start;
            if (length == 1) return text[start] == (byte)'1';
            if (length != 4) return false;
            return (text[start] | 0x20) == (byte)'t'
                && (text[start + 1] | 0x20) == (byte)'r'
                && (text[start + 2] | 0x20) == (byte)'u'
                && (text[start + 3] | 0x20) == (byte)'e';
        }

        private static bool IsXmlWhitespace(char value) => value is ' ' or '\t' or '\r' or '\n';
    }
}
