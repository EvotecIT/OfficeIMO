namespace OfficeIMO.Excel {
    /// <summary>
    /// Recognizes relative OPC paths that already need no URI decoding, escaping,
    /// or dot-segment normalization. Other spellings use the full URI validation.
    /// </summary>
    internal static class OpenXmlPartName {
        internal static bool IsCanonicalPath(string value, int startIndex = 0) {
            int segmentStart = startIndex;
            for (int index = startIndex; index < value.Length; index++) {
                char character = value[index];
                if (character == '/') {
                    if (!IsOrdinarySegment(value, segmentStart, index)) return false;
                    segmentStart = index + 1;
                } else if (!(character is >= 'a' and <= 'z'
                    or >= 'A' and <= 'Z'
                    or >= '0' and <= '9'
                    or '-' or '.' or '_' or '~')) {
                    return false;
                }
            }

            return IsOrdinarySegment(value, segmentStart, value.Length);
        }

        private static bool IsOrdinarySegment(string value, int start, int end) {
            int length = end - start;
            return length > 0 && !(value[start] == '.'
                && (length == 1 || length == 2 && value[start + 1] == '.'));
        }
    }
}
