namespace OfficeIMO.Markdown;

public sealed partial class MarkdownNativeDocument {
    private string FormatDefinitionListDefinitionReplacement(
        MarkdownNativeDefinitionListDefinition definition,
        string replacementMarkdown) {
        replacementMarkdown ??= string.Empty;
        if (!definition.SourceSpan.HasValue || replacementMarkdown.IndexOfAny(new[] { '\r', '\n' }) < 0) {
            return replacementMarkdown;
        }

        var continuationIndent = DetectDefinitionContinuationIndent(definition.SourceSpan.Value);
        if (continuationIndent.Length == 0) {
            return replacementMarkdown;
        }

        var normalized = replacementMarkdown
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
        var lines = normalized.Split('\n');
        for (var i = 1; i < lines.Length; i++) {
            if (lines[i].Length == 0) {
                continue;
            }

            lines[i] = continuationIndent + lines[i];
        }

        return string.Join("\n", lines);
    }

    private string DetectDefinitionContinuationIndent(MarkdownSourceSpan bodySpan) {
        var existingIndent = FindExistingDefinitionContinuationIndent(bodySpan);
        if (existingIndent != null) {
            return existingIndent;
        }

        if (!bodySpan.StartColumn.HasValue || bodySpan.StartColumn.Value <= 1) {
            return string.Empty;
        }

        var firstLinePrefix = GetSourceLinePrefix(bodySpan.StartLine, bodySpan.StartColumn.Value);
        return IsMarkerDefinitionPrefix(firstLinePrefix)
            ? new string(' ', bodySpan.StartColumn.Value - 1)
            : "  ";
    }

    private string? FindExistingDefinitionContinuationIndent(MarkdownSourceSpan bodySpan) {
        for (var lineNumber = bodySpan.StartLine + 1; lineNumber <= bodySpan.EndLine; lineNumber++) {
            if (!TryGetSourceLine(lineNumber, out var line) || string.IsNullOrWhiteSpace(line)) {
                continue;
            }

            var indentLength = 0;
            while (indentLength < line.Length && char.IsWhiteSpace(line[indentLength])) {
                indentLength++;
            }

            if (indentLength == 0) {
                continue;
            }

            return line.Substring(0, indentLength);
        }

        return null;
    }

    private string GetSourceLinePrefix(int lineNumber, int startColumn) {
        if (!TryGetSourceLine(lineNumber, out var line) || line.Length == 0) {
            return string.Empty;
        }

        var prefixLength = Math.Min(line.Length, Math.Max(0, startColumn - 1));
        return line.Substring(0, prefixLength);
    }

    private bool TryGetSourceLine(int lineNumber, out string line) {
        line = string.Empty;
        if (!TryGetLineStartOffset(lineNumber, out var lineStart)) {
            return false;
        }

        var lineEndExclusive = SourceMarkdown.Length;
        for (var i = lineStart; i < SourceMarkdown.Length; i++) {
            if (SourceMarkdown[i] == '\n') {
                lineEndExclusive = i;
                break;
            }
        }

        if (lineEndExclusive > lineStart && SourceMarkdown[lineEndExclusive - 1] == '\r') {
            lineEndExclusive--;
        }

        line = SourceMarkdown.Substring(lineStart, lineEndExclusive - lineStart);
        return true;
    }

    private static bool IsMarkerDefinitionPrefix(string prefix) {
        if (string.IsNullOrEmpty(prefix)) {
            return false;
        }

        var colonIndex = prefix.IndexOf(':');
        if (colonIndex < 0) {
            return false;
        }

        for (var i = 0; i < colonIndex; i++) {
            if (prefix[i] != ' ') {
                return false;
            }
        }

        for (var i = colonIndex + 1; i < prefix.Length; i++) {
            if (prefix[i] != ' ' && prefix[i] != '\t') {
                return false;
            }
        }

        return true;
    }
}
