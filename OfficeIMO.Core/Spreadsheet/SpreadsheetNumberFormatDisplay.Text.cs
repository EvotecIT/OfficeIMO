using System.Globalization;
using System.Text;

namespace OfficeIMO.Spreadsheet;

internal static partial class SpreadsheetNumberFormatDisplay {
    /// <summary>Applies the text section without interpreting the value as a number or dropping unknown syntax.</summary>
    internal static string FormatTextValue(string value, string? formatCode) {
        if (string.IsNullOrEmpty(formatCode) || string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase)
            || string.Equals(formatCode, "@", StringComparison.Ordinal)) return value;
        var builder = new StringBuilder();
        ComposeTextValue(value, formatCode, (text, _) => builder.Append(text));
        return builder.ToString();
    }

    /// <summary>Emits display literals separately from value placeholders so rich-text consumers retain run metrics.</summary>
    internal static void ComposeTextValue(string value, string? formatCode, Action<string, bool> append) {
        if (string.IsNullOrEmpty(formatCode)) { append(value, true); return; }
        string[] sections = SplitNumberFormatSections(formatCode!);
        string section = sections.Length == 4 ? sections[3] : sections.Length == 1 ? sections[0] : string.Empty;
        if (string.Equals(section, "General", StringComparison.OrdinalIgnoreCase)) { append(value, true); return; }
        var syntax = SpreadsheetNumberFormatSyntax.Parse(section);
        if (!syntax.IsValid || sections.Length != 4
            && !syntax.Tokens.Any(token => token.Kind == SpreadsheetNumberFormatTokenKind.TextPlaceholder)) {
            append(value, true);
            return;
        }
        foreach (var token in syntax.Tokens) {
            switch (token.Kind) {
                case SpreadsheetNumberFormatTokenKind.TextPlaceholder: append(value, true); break;
                case SpreadsheetNumberFormatTokenKind.Literal:
                    if (token.Text[0] == '_') append(" ", false);
                    else if (token.Text[0] != '*') append(token.Value, false);
                    break;
                case SpreadsheetNumberFormatTokenKind.Currency: append(token.Value, false); break;
                case SpreadsheetNumberFormatTokenKind.BracketedDirective:
                    // Color/locale directives have no textual representation. Keep unknown
                    // directives intact rather than silently deleting imported syntax.
                    string directive = token.Value;
                    bool indexedColor = directive.StartsWith("Color", StringComparison.OrdinalIgnoreCase)
                        && int.TryParse(directive.Substring(5), NumberStyles.None, CultureInfo.InvariantCulture, out int color)
                        && color >= 1 && color <= 56;
                    if (!directive.StartsWith("$", StringComparison.Ordinal)
                        && !new[] { "Black", "Blue", "Cyan", "Green", "Magenta", "Red", "White", "Yellow" }
                            .Contains(directive, StringComparer.OrdinalIgnoreCase)
                        && !indexedColor) append(token.Text, false);
                    break;
                default: append(token.Text, false); break;
            }
        }
    }
}
