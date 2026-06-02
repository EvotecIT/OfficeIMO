using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static bool IsDefinitionLine(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        var trimmed = line.TrimStart();
        if (IsAtxHeading(trimmed, out _, out _)) return false; // headings take priority over definition lists
        if (IsUnorderedListLine(trimmed, out _, out _, out _)) return false; // list items with ":" are not definition terms
        if (IsOrderedListLine(trimmed, out _, out _)) return false; // numbered list items with ":" are not definition terms
        if (StartsWithReferenceDefinitionLikeLabel(trimmed)) return false; // malformed or valid link ref definitions should not become <dl>
        return TryGetDefinitionSeparator(line, out _);
    }

    private static bool ShouldTreatAsDefinitionLine(IReadOnlyList<string>? lines, int index, MarkdownReaderOptions options) {
        if (lines == null || index < 0 || index >= lines.Count) return false;
        if (options == null || !options.DefinitionLists) return false;

        var line = lines[index] ?? string.Empty;
        if (!IsDefinitionLineBlockCandidate(line)) return false;
        if (!options.PreferNarrativeSingleLineDefinitions) return true;

        return HasAdjacentDefinitionLine(lines, index) || HasDefinitionContinuation(lines, index);
    }

    private static bool HasAdjacentDefinitionLine(IReadOnlyList<string> lines, int index) {
        return IsDefinitionLineBlockCandidate(index > 0 ? lines[index - 1] : null)
               || IsDefinitionLineBlockCandidate(index + 1 < lines.Count ? lines[index + 1] : null);
    }

    private static bool HasDefinitionContinuation(IReadOnlyList<string> lines, int index) {
        if (lines == null || index < 0 || index >= lines.Count) {
            return false;
        }

        var line = lines[index] ?? string.Empty;
        int continuationIndent = CountLeadingIndentColumns(line) + 2;
        for (int i = index + 1; i < lines.Count; i++) {
            var next = lines[i] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(next)) {
                continue;
            }

            return CountLeadingIndentColumns(next) >= continuationIndent;
        }

        return false;
    }

    private static bool IsDefinitionLineBlockCandidate(string? line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        string safeLine = line!;

        int leading = 0;
        while (leading < safeLine.Length && safeLine[leading] == ' ') leading++;
        if (leading >= 4) return false;
        if (leading < safeLine.Length && safeLine[leading] == '\t') return false;

        return IsDefinitionLine(safeLine);
    }

    private static bool TryGetDefinitionSeparator(string line, out int idx) {
        idx = -1;
        if (string.IsNullOrWhiteSpace(line)) return false;
        int start = 0;
        while (start < line.Length) {
            int pos = line.IndexOf(':', start);
            if (pos < 0) return false;
            if (pos > 0 && pos + 1 < line.Length && line[pos + 1] == ' ') {
                var term = line.Substring(0, pos).Trim();
                if (LooksLikeDefinitionTerm(term)) {
                    idx = pos;
                    return true;
                }
            }
            start = pos + 1;
        }
        return false;
    }

    private static bool LooksLikeDefinitionTerm(string term) {
        if (string.IsNullOrWhiteSpace(term)) return false;
        return !ContainsLiteralAutolinkLikeToken(term);
    }

    private static bool ContainsLiteralAutolinkLikeToken(string text) {
        foreach (var rawToken in text.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries)) {
            if (LooksLikeMarkdownLinkToken(rawToken)) continue;

            var token = rawToken
                .TrimStart('(', '[', '{', '"', '\'')
                .TrimEnd(')', ']', '}', '"', '\'', '.', ',', ';', '!', '?');
            if (string.IsNullOrWhiteSpace(token)) continue;

            if (token[0] == '<' &&
                TryParseAngleAutolink(token, 0, out int angleConsumed, out _, out _) &&
                angleConsumed == token.Length) {
                return true;
            }

            if ((token[0] == 'h' || token[0] == 'H') &&
                StartsWithHttp(token, 0, out int httpEnd) &&
                httpEnd == token.Length) {
                return true;
            }

            if ((token[0] == 'w' || token[0] == 'W') &&
                StartsWithWww(token, 0, out int wwwEnd) &&
                wwwEnd == token.Length) {
                return true;
            }

            if (IsEmailStartChar(token[0]) &&
                TryConsumePlainEmail(token, 0, out int emailEnd, out _) &&
                emailEnd == token.Length) {
                return true;
            }
        }

        return false;
    }

    private static bool LooksLikeMarkdownLinkToken(string token) {
        if (string.IsNullOrWhiteSpace(token)) return false;

        int start = token[0] == '!' ? 1 : 0;
        if (start >= token.Length || token[start] != '[') return false;

        int closeLabel = token.IndexOf(']', start + 1);
        if (closeLabel < 0 || closeLabel + 1 >= token.Length) return false;

        return (token[closeLabel + 1] == '(' && token[token.Length - 1] == ')') ||
               (token[closeLabel + 1] == '[' && token[token.Length - 1] == ']');
    }
}
