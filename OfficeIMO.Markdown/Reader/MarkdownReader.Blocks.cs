using System.Globalization;
using System.Linq;

namespace OfficeIMO.Markdown;

/// <summary>
/// Block parsing helpers for <see cref="MarkdownReader"/>.
/// </summary>
public static partial class MarkdownReader {
    private static bool IsAtxHeading(string line, out int level, out string text) {
        return TryGetAtxHeadingContentRange(line, out level, out _, out _, out text);
    }

    private static bool TryGetAtxHeadingContentRange(string line, out int level, out int contentStart, out int contentEnd, out string text) {
        level = 0;
        contentStart = 0;
        contentEnd = 0;
        text = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;

        int indent = 0;
        while (indent < line.Length && indent < 4 && line[indent] == ' ') indent++;
        if (indent >= 4) return false;

        int i = indent;
        while (i < line.Length && line[i] == '#') i++;

        int count = i - indent;
        if (count < 1 || count > 6) return false;
        if (i < line.Length && !char.IsWhiteSpace(line[i])) return false;

        contentStart = i;
        while (contentStart < line.Length && char.IsWhiteSpace(line[contentStart])) contentStart++;
        if (contentStart >= line.Length) {
            level = count;
            text = string.Empty;
            contentEnd = contentStart;
            return true;
        }

        contentEnd = line.Length;
        while (contentEnd > contentStart && char.IsWhiteSpace(line[contentEnd - 1])) contentEnd--;

        int closingStart = contentEnd;
        while (closingStart > contentStart && line[closingStart - 1] == '#') closingStart--;
        if (closingStart < contentEnd) {
            int beforeClosing = closingStart - 1;
            if (beforeClosing >= contentStart && char.IsWhiteSpace(line[beforeClosing])) {
                contentEnd = beforeClosing;
                while (contentEnd > contentStart && char.IsWhiteSpace(line[contentEnd - 1])) contentEnd--;
            }
        }

        level = count;
        text = line.Substring(contentStart, contentEnd - contentStart);
        return true;
    }

}
