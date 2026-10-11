namespace OfficeIMO.Reader {

    /// <summary>Escapes adapter link destinations without changing their URI semantics.</summary>
    internal static class ReaderMarkdownDestination {
        internal static string EscapeDestination(string value) {
            var escaped = new System.Text.StringBuilder(value.Length);
            foreach (char character in value) {
                if (char.IsWhiteSpace(character)) {
                    foreach (byte utf8Byte in System.Text.Encoding.UTF8.GetBytes(character.ToString())) {
                        escaped.Append('%').Append(utf8Byte.ToString("X2"));
                    }
                } else if (character == '\\' || character == '(' || character == ')') {
                    escaped.Append('\\').Append(character);
                } else {
                    escaped.Append(character);
                }
            }
            return escaped.ToString();
        }
    }
}
