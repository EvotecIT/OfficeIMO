using System;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word;

public static partial class WordPaperText {
    private const string IdentifierBoundaries = @"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])|(?<=[A-Za-z])(?=[0-9])|(?<=[0-9])(?=[A-Za-z])";
    private static readonly string WordIdentifierBreaks = IdentifierBoundaries.Replace("(?=", "\u200C(?=");

    /// <summary>Adds Word's invisible U+200C no-width optional break at case and digit boundaries without changing the visible identifier.</summary>
    /// <param name="value">The exact source identifier.</param>
    /// <returns>Display text containing optional breaks. Use <see cref="RemoveOptionalBreaks"/> for plain-text targets.</returns>
    public static string WithIdentifierBreaks(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        return Regex.Replace(value, IdentifierBoundaries, "\u200C");
    }

    /// <summary>Adds invisible U+200B optional breaks after selected separators, preserving whole identifier labels.</summary>
    /// <param name="value">The exact source identifier, such as a hostname, path or record value.</param>
    /// <param name="separators">Characters after which a line may wrap; for hostnames, supply dots and hyphens.</param>
    /// <returns>Display text with no visible spaces added. Repeated application is idempotent. Use <see cref="RemoveOptionalBreaks"/> for plain text.</returns>
    /// <remarks>This overload does not insert breaks at case or digit boundaries inside labels.</remarks>
    public static string WithIdentifierBreaks(string value, params char[] separators) {
        if (value == null) { throw new ArgumentNullException(nameof(value)); }
        if (separators == null) { throw new ArgumentNullException(nameof(separators)); }
        if (value.Length == 0 || separators.Length == 0) { return value; }
        var text = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            text.Append(character);
            if (index + 1 < value.Length && value[index + 1] != '\u200B' && Array.IndexOf(separators, character) >= 0) {
                text.Append('\u200B');
            }
        }
        return text.ToString();
    }

    /// <summary>Removes authored identifier breaks and portable U+200B markers for copy, search and plain-text export.</summary>
    /// <param name="value">Extracted text, optionally containing invisible breaks.</param>
    /// <returns>The original text without optional break markers; visible spacing and linguistic join controls outside the authored Latin case/digit boundaries are preserved.</returns>
    public static string RemoveOptionalBreaks(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        return Regex.Replace(value, WordIdentifierBreaks, string.Empty).Replace("\u200B", string.Empty);
    }

    // Word's optional break uses a join-control code point. Convert only the Latin
    // identifier boundaries authored above, leaving linguistic joining text intact.
    internal static string NormalizeIdentifierBreaksForLayout(string value) => value.IndexOf('\u200C') < 0
        ? value : Regex.Replace(value, WordIdentifierBreaks, "\u200B");
}
