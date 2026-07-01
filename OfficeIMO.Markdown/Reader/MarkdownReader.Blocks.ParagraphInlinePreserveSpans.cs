namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private readonly struct ParagraphInlinePreserveRange {
        public ParagraphInlinePreserveRange(int contentStart, int contentEnd) {
            ContentStart = contentStart;
            ContentEnd = contentEnd;
        }

        public int ContentStart { get; }
        public int ContentEnd { get; }

        public bool ContainsOffset(int offset) =>
            offset >= ContentStart && offset < ContentEnd;
    }

    private static bool[] FindParagraphLineBreaksInsideMatchedInlinePreserveSpans(IReadOnlyList<string> lines, MarkdownReaderOptions options) {
        if (lines == null || lines.Count <= 1) {
            return Array.Empty<bool>();
        }

        var text = string.Join("\n", lines);
        if (text.IndexOf('`') < 0 && text.IndexOf('<') < 0) {
            return new bool[Math.Max(0, lines.Count - 1)];
        }

        var ranges = FindMatchedCodeSpanRanges(text);
        if (options?.InlineHtml == true) {
            ranges.AddRange(FindMatchedRawInlineHtmlTagRanges(text));
        }

        var result = new bool[lines.Count - 1];
        var offset = 0;
        for (var lineIndex = 0; lineIndex < lines.Count - 1; lineIndex++) {
            offset += lines[lineIndex]?.Length ?? 0;
            result[lineIndex] = ranges.Any(range => range.ContainsOffset(offset));
            offset++;
        }

        return result;
    }

    private static List<ParagraphInlinePreserveRange> FindMatchedCodeSpanRanges(string text) {
        var ranges = new List<ParagraphInlinePreserveRange>();
        for (var position = 0; position < text.Length;) {
            if (text[position] != '`') {
                position++;
                continue;
            }

            var openingStart = position;
            var fenceLength = CountBacktickRun(text, position);
            var search = openingStart + fenceLength;
            var closingStart = -1;
            while (search < text.Length) {
                if (text[search] != '`') {
                    search++;
                    continue;
                }

                var candidateLength = CountBacktickRun(text, search);
                if (candidateLength == fenceLength) {
                    closingStart = search;
                    break;
                }

                search += candidateLength;
            }

            if (closingStart < 0) {
                position = openingStart + fenceLength;
                continue;
            }

            ranges.Add(new ParagraphInlinePreserveRange(openingStart + fenceLength, closingStart));
            position = closingStart + fenceLength;
        }

        return ranges;
    }

    private static List<ParagraphInlinePreserveRange> FindMatchedRawInlineHtmlTagRanges(string text) {
        var ranges = new List<ParagraphInlinePreserveRange>();
        for (int position = 0; position < text.Length;) {
            int tagStart = text.IndexOf('<', position);
            if (tagStart < 0) {
                break;
            }

            if (!HtmlBlockParser.TryParseHtmlTag(text.Substring(tagStart), out _, out _, out int endIndex)) {
                position = tagStart + 1;
                continue;
            }

            ranges.Add(new ParagraphInlinePreserveRange(tagStart + 1, tagStart + endIndex));
            position = tagStart + endIndex + 1;
        }

        return ranges;
    }

    private static int CountBacktickRun(string text, int start) {
        var length = 0;
        while (start + length < text.Length && text[start + length] == '`') {
            length++;
        }

        return length;
    }
}
