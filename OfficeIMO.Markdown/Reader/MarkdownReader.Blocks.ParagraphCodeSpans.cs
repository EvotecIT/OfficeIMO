namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private readonly struct ParagraphCodeSpanRange {
        public ParagraphCodeSpanRange(int contentStart, int closingStart) {
            ContentStart = contentStart;
            ClosingStart = closingStart;
        }

        public int ContentStart { get; }
        public int ClosingStart { get; }

        public bool ContainsOffset(int offset) =>
            offset >= ContentStart && offset < ClosingStart;
    }

    private static bool[] FindParagraphLineBreaksInsideMatchedCodeSpans(IReadOnlyList<string> lines) {
        if (lines == null || lines.Count <= 1) {
            return Array.Empty<bool>();
        }

        var text = string.Join("\n", lines);
        if (text.IndexOf('`') < 0) {
            return new bool[Math.Max(0, lines.Count - 1)];
        }

        var codeSpans = FindMatchedCodeSpanRanges(text);
        var result = new bool[lines.Count - 1];
        var offset = 0;
        for (var lineIndex = 0; lineIndex < lines.Count - 1; lineIndex++) {
            offset += lines[lineIndex]?.Length ?? 0;
            result[lineIndex] = codeSpans.Any(range => range.ContainsOffset(offset));
            offset++;
        }

        return result;
    }

    private static List<ParagraphCodeSpanRange> FindMatchedCodeSpanRanges(string text) {
        var ranges = new List<ParagraphCodeSpanRange>();
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

            ranges.Add(new ParagraphCodeSpanRange(openingStart + fenceLength, closingStart));
            position = closingStart + fenceLength;
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
