using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static bool TryParseReferenceDestinationAndMultilineTitle(
        string[] lines,
        int destinationLineIndex,
        string rest,
        int contentStartColumnZeroBased,
        MarkdownReaderState? state,
        out string url,
        out string? title,
        out MarkdownSourceSpan? urlSpan,
        out MarkdownSourceSpan? titleSpan,
        out int titleEndLineIndex) {
        url = string.Empty;
        title = null;
        urlSpan = null;
        titleSpan = null;
        titleEndLineIndex = destinationLineIndex;

        if (string.IsNullOrEmpty(rest) || lines == null || destinationLineIndex < 0 || destinationLineIndex >= lines.Length) {
            return false;
        }

        int start = 0;
        while (start < rest.Length && IsLinkWhitespace(rest[start])) {
            start++;
        }

        if (start >= rest.Length || rest[start] == '<') {
            return false;
        }

        int destinationEnd = start;
        while (destinationEnd < rest.Length && !IsLinkWhitespace(rest[destinationEnd])) {
            destinationEnd++;
        }

        if (destinationEnd <= start || destinationEnd >= rest.Length) {
            return false;
        }

        int titleStart = destinationEnd;
        while (titleStart < rest.Length && IsLinkWhitespace(rest[titleStart])) {
            titleStart++;
        }

        if (titleStart >= rest.Length) {
            return false;
        }

        char opener = rest[titleStart];
        char closer = opener switch {
            '"' => '"',
            '\'' => '\'',
            '(' => ')',
            _ => '\0'
        };

        if (closer == '\0') {
            return false;
        }

        var builder = new StringBuilder();
        string firstTitleContent = rest.Substring(titleStart + 1);
        if (ContainsUnescapedTitleDelimiter(firstTitleContent, 0, firstTitleContent.Length, closer)) {
            return false;
        }

        builder.Append(firstTitleContent);

        for (int lineIndex = destinationLineIndex + 1; lineIndex < lines.Length; lineIndex++) {
            string line = lines[lineIndex] ?? string.Empty;
            int trimmedEndExclusive = line.Length;
            while (trimmedEndExclusive > 0 && IsLinkWhitespace(line[trimmedEndExclusive - 1])) {
                trimmedEndExclusive--;
            }

            if (trimmedEndExclusive == 0) {
                return false;
            }

            int closerIndex = trimmedEndExclusive - 1;
            if (line[closerIndex] == closer) {
                if (ContainsUnescapedTitleDelimiter(line, 0, closerIndex, closer)) {
                    return false;
                }

                builder.Append('\n');
                builder.Append(line.Substring(0, closerIndex));
                url = DecodeLinkDestinationOrTitle(rest.Substring(start, destinationEnd - start));
                title = DecodeLinkDestinationOrTitle(builder.ToString());
                urlSpan = CreateSpan(
                    state,
                    state?.SourceLineOffset + destinationLineIndex + 1 ?? destinationLineIndex + 1,
                    contentStartColumnZeroBased + start + 1,
                    state?.SourceLineOffset + destinationLineIndex + 1 ?? destinationLineIndex + 1,
                    contentStartColumnZeroBased + destinationEnd);
                titleSpan = CreateSpan(
                    state,
                    state?.SourceLineOffset + destinationLineIndex + 1 ?? destinationLineIndex + 1,
                    contentStartColumnZeroBased + titleStart + 2,
                    state?.SourceLineOffset + lineIndex + 1 ?? lineIndex + 1,
                    Math.Max(1, closerIndex));
                titleEndLineIndex = lineIndex;
                return true;
            }

            if (ContainsUnescapedTitleDelimiter(line, 0, trimmedEndExclusive, closer)) {
                return false;
            }

            builder.Append('\n');
            builder.Append(line);
        }

        return false;
    }
}
