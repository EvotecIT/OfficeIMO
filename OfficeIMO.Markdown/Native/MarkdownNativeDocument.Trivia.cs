namespace OfficeIMO.Markdown;

public sealed partial class MarkdownNativeDocument {
    /// <summary>Enumerates document-level source trivia of the requested kind in source order.</summary>
    public IEnumerable<MarkdownNativeSourceTrivia> EnumerateSourceTrivia(MarkdownNativeSourceTriviaKind kind) {
        for (var i = 0; i < SourceTrivia.Count; i++) {
            if (SourceTrivia[i].Kind == kind) {
                yield return SourceTrivia[i];
            }
        }
    }

    /// <summary>Finds the first document-level source trivia whose span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeSourceTrivia? FindSourceTriviaAtPosition(int lineNumber, int columnNumber) {
        for (var i = 0; i < SourceTrivia.Count; i++) {
            if (SourceTrivia[i].SourceSpan.ContainsPosition(lineNumber, columnNumber)) {
                return SourceTrivia[i];
            }
        }

        return null;
    }

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
                trivia.Add(CreateTrivia(
                    MarkdownNativeSourceTriviaKind.BlankLine,
                    sourceMarkdown,
                    lineNumber,
                    lineStartOffset,
                    lineLength));
            } else {
                var leadingLength = CountLeadingHorizontalWhitespace(sourceMarkdown, lineStartOffset, lineLength);
                if (leadingLength > 0) {
                    trivia.Add(CreateTrivia(
                        MarkdownNativeSourceTriviaKind.LeadingWhitespace,
                        sourceMarkdown,
                        lineNumber,
                        lineStartOffset,
                        leadingLength));
                }

                var trailingLength = CountTrailingHorizontalWhitespace(sourceMarkdown, lineStartOffset, lineLength);
                if (trailingLength > 0) {
                    trivia.Add(CreateTrivia(
                        MarkdownNativeSourceTriviaKind.TrailingWhitespace,
                        sourceMarkdown,
                        lineNumber,
                        lineBreakOffset - trailingLength,
                        trailingLength));
                }
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

    private static MarkdownNativeSourceTrivia CreateTrivia(
        MarkdownNativeSourceTriviaKind kind,
        string sourceMarkdown,
        int lineNumber,
        int startOffset,
        int length) {
        var text = length == 0
            ? string.Empty
            : sourceMarkdown.Substring(startOffset, length);
        var endOffsetInclusive = length == 0
            ? startOffset - 1
            : startOffset + length - 1;
        var startColumn = GetColumnNumber(sourceMarkdown, startOffset);
        var endColumn = length == 0
            ? startColumn
            : startColumn + length - 1;
        var sourceSpan = new MarkdownSourceSpan(
            lineNumber,
            startColumn,
            lineNumber,
            endColumn,
            startOffset,
            endOffsetInclusive);

        return new MarkdownNativeSourceTrivia(kind, text, sourceSpan);
    }

    private static bool IsBlankLine(string sourceMarkdown, int lineStartOffset, int lineLength) {
        for (var i = 0; i < lineLength; i++) {
            if (!char.IsWhiteSpace(sourceMarkdown[lineStartOffset + i])) {
                return false;
            }
        }

        return true;
    }

    private static int CountLeadingHorizontalWhitespace(string sourceMarkdown, int lineStartOffset, int lineLength) {
        var count = 0;
        while (count < lineLength && IsHorizontalWhitespace(sourceMarkdown[lineStartOffset + count])) {
            count++;
        }

        return count;
    }

    private static int CountTrailingHorizontalWhitespace(string sourceMarkdown, int lineStartOffset, int lineLength) {
        var count = 0;
        while (count < lineLength && IsHorizontalWhitespace(sourceMarkdown[lineStartOffset + lineLength - count - 1])) {
            count++;
        }

        return count;
    }

    private static bool IsHorizontalWhitespace(char value) => value == ' ' || value == '\t';

    private static int GetColumnNumber(string sourceMarkdown, int offset) {
        var column = 1;
        for (var i = offset - 1; i >= 0; i--) {
            if (IsLineBreakStart(sourceMarkdown, i, out _)) {
                break;
            }

            column++;
        }

        return column;
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
