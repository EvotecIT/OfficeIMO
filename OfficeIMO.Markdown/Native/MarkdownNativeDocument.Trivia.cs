namespace OfficeIMO.Markdown;

public sealed partial class MarkdownNativeDocument {
    private static IReadOnlyList<MarkdownNativeSourceTrivia> CreateSourceTrivia(string? sourceMarkdown) {
        sourceMarkdown ??= string.Empty;
        if (sourceMarkdown.Length == 0) {
            return Array.Empty<MarkdownNativeSourceTrivia>();
        }

        var trivia = new List<MarkdownNativeSourceTrivia>();
        var lineNumber = 1;
        var lineStartOffset = 0;
        while (lineStartOffset < sourceMarkdown.Length) {
            var lineBreakOffset = lineStartOffset;
            while (lineBreakOffset < sourceMarkdown.Length && !IsLineBreakStart(sourceMarkdown, lineBreakOffset, out _)) {
                lineBreakOffset++;
            }

            var lineLength = lineBreakOffset - lineStartOffset;
            if (IsBlankLine(sourceMarkdown, lineStartOffset, lineLength)) {
                var lineText = lineLength == 0
                    ? string.Empty
                    : sourceMarkdown.Substring(lineStartOffset, lineLength);
                var endOffsetInclusive = lineLength == 0
                    ? lineStartOffset - 1
                    : lineBreakOffset - 1;
                var sourceSpan = new MarkdownSourceSpan(
                    lineNumber,
                    1,
                    lineNumber,
                    Math.Max(1, lineLength),
                    lineStartOffset,
                    endOffsetInclusive);
                trivia.Add(new MarkdownNativeSourceTrivia(MarkdownNativeSourceTriviaKind.BlankLine, lineText, sourceSpan));
            }

            if (lineBreakOffset >= sourceMarkdown.Length) {
                break;
            }

            IsLineBreakStart(sourceMarkdown, lineBreakOffset, out var lineBreakLength);
            lineStartOffset = lineBreakOffset + lineBreakLength;
            lineNumber++;
        }

        return trivia.Count == 0 ? Array.Empty<MarkdownNativeSourceTrivia>() : trivia;
    }

    private static bool IsBlankLine(string sourceMarkdown, int lineStartOffset, int lineLength) {
        for (var i = 0; i < lineLength; i++) {
            if (!char.IsWhiteSpace(sourceMarkdown[lineStartOffset + i])) {
                return false;
            }
        }

        return true;
    }

    private static bool IsLineBreakStart(string sourceMarkdown, int offset, out int length) {
        if (sourceMarkdown[offset] == '\r') {
            length = offset + 1 < sourceMarkdown.Length && sourceMarkdown[offset + 1] == '\n'
                ? 2
                : 1;
            return true;
        }

        if (sourceMarkdown[offset] == '\n') {
            length = 1;
            return true;
        }

        length = 0;
        return false;
    }
}
