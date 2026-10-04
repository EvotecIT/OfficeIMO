namespace OfficeIMO.Reader;

/// <summary>Literal-text escaping for adapters that project source text without depending on the Markdown document engine.</summary>
internal static class ReaderMarkdownEscaping {
    // CommonMark permits backslash escapes for ASCII punctuation. Escaping all of it
    // also protects block starts, HTML, links, table delimiters and emphasis extensions.
    internal static bool IsLiteralPunctuation(char value) =>
        value is >= '!' and <= '/' or >= ':' and <= '@' or >= '[' and <= '`' or >= '{' and <= '~';
}
