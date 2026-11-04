namespace OfficeIMO.Markdown;

using System.Text;

internal static partial class MarkdownEscaper {
    internal static bool UsesBackslashLiteralText => MarkdownRenderContext.Options?.TextEscapingMode == MarkdownTextEscapingMode.LiteralBackslash;

    internal static string EscapeTextForWriter(string? text) => UsesBackslashLiteralText ? EscapeBackslashLiteralText(text) : EscapeText(text);

    internal static string EscapeLiteralTextForWriter(string? text) => UsesBackslashLiteralText ? EscapeBackslashLiteralText(text) : EscapeLiteralText(text);

    internal static string EscapeLineStartTextForWriter(string? text) => UsesBackslashLiteralText ? EscapeBackslashLiteralText(text) : EscapeTextAndLineStarts(text);

    /// <summary>
    /// Escapes literal content once, using Markdown punctuation rather than character references. The existing
    /// line-start rules protect headings, lists and fences; explicit inline/structured nodes retain their markup.
    /// </summary>
    internal static string EscapeBackslashLiteralText(string? text) {
        if (string.IsNullOrEmpty(text)) { return string.Empty; }
        var escaped = new StringBuilder(text!.Length);
        bool superscriptPair = text.IndexOf('^') != text.LastIndexOf('^');
        for (int index = 0; index < text.Length; index++) {
            char character = text[index];
            bool entity = character == '&' && IsEntityLikeAmpersand(text, index);
            bool highlight = character == '=' && ((index > 0 && text[index - 1] == '=') || (index + 1 < text.Length && text[index + 1] == '='));
            bool insertion = character == '+' && ((index > 0 && text[index - 1] == '+') || (index + 1 < text.Length && text[index + 1] == '+'));
            if ("<>\\`*_{}[]|#~".IndexOf(character) >= 0 || entity || highlight || insertion || (character == '^' && superscriptPair)) { escaped.Append('\\'); }
            escaped.Append(character);
        }
        return EscapeMarkdownLineStarts(escaped.ToString(), preserveDefinitionText: true);
    }
}
