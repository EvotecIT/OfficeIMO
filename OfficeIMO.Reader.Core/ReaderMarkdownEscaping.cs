namespace OfficeIMO.Reader;

/// <summary>Literal-text escaping for adapters that project source text without depending on the Markdown document engine.</summary>
internal static class ReaderMarkdownEscaping {
    /// <summary>Escapes literal source text without activating Markdown markup, links or raw HTML.</summary>
    internal static string EscapeLiteral(string value, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var escaped = new System.Text.StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if ((index & 511) == 0) cancellationToken.ThrowIfCancellationRequested();
            char character = value[index];
            if (IsLiteralPunctuation(character)) escaped.Append('\\');
            escaped.Append(character);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return escaped.ToString();
    }

    // CommonMark permits backslash escapes for ASCII punctuation. Escaping all of it
    // also protects block starts, HTML, links, table delimiters and emphasis extensions.
    internal static bool IsLiteralPunctuation(char value) =>
        value is >= '!' and <= '/' or >= ':' and <= '@' or >= '[' and <= '`' or >= '{' and <= '~';
}
