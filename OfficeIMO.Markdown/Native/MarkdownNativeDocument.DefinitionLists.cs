namespace OfficeIMO.Markdown;

public sealed partial class MarkdownNativeDocument {
    private string FormatBlockSourceFieldReplacement(
        MarkdownNativeBlockSourceField field,
        string replacementMarkdown) {
        if (field != null &&
            string.Equals(field.Name, "definitionBody", StringComparison.OrdinalIgnoreCase) &&
            TryGetDefinitionListDefinition(field, out var definition)) {
            return FormatDefinitionListDefinitionReplacement(definition, replacementMarkdown);
        }

        return replacementMarkdown ?? string.Empty;
    }

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
        var lineEnding = replacementMarkdown.IndexOf("\r\n", StringComparison.Ordinal) >= 0
            ? "\r\n"
            : "\n";
        var lines = normalized.Split('\n');
        for (var i = 1; i < lines.Length; i++) {
            if (lines[i].Length == 0 || char.IsWhiteSpace(lines[i][0])) {
                continue;
            }

            lines[i] = continuationIndent + lines[i];
        }

        return string.Join(lineEnding, lines);
    }

    private static bool TryGetDefinitionListDefinition(
        MarkdownNativeBlockSourceField field,
        out MarkdownNativeDefinitionListDefinition definition) {
        definition = null!;
        if (field == null ||
            field.Index < 0 ||
            field.Block is not MarkdownNativeDefinitionListBlock definitionList) {
            return false;
        }

        var currentIndex = 0;
        for (var groupIndex = 0; groupIndex < definitionList.Groups.Count; groupIndex++) {
            var definitions = definitionList.Groups[groupIndex].Definitions;
            for (var definitionIndex = 0; definitionIndex < definitions.Count; definitionIndex++) {
                if (currentIndex == field.Index) {
                    definition = definitions[definitionIndex];
                    return true;
                }

                currentIndex++;
            }
        }

        return false;
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
        if (!TryGetSourceLineBounds(lineNumber, out var lineStart, out var lineEndExclusive) || lineEndExclusive <= lineStart) {
            return string.Empty;
        }

        var prefixEnd = MarkdownSourceColumns.ResolveVisualColumnStartOffset(
            SourceMarkdown,
            lineStart,
            lineEndExclusive,
            startColumn);
        return SourceMarkdown.Substring(lineStart, prefixEnd - lineStart);
    }

    private bool TryGetSourceLine(int lineNumber, out string line) {
        line = string.Empty;
        if (!TryGetSourceLineBounds(lineNumber, out var lineStart, out var lineEndExclusive)) {
            return false;
        }

        line = SourceMarkdown.Substring(lineStart, lineEndExclusive - lineStart);
        return true;
    }

    private bool TryGetSourceLineBounds(int lineNumber, out int lineStart, out int lineEndExclusive) {
        lineStart = 0;
        lineEndExclusive = 0;
        if (!TryGetLineStartOffset(lineNumber, out lineStart)) {
            return false;
        }

        lineEndExclusive = SourceMarkdown.Length;
        for (var i = lineStart; i < SourceMarkdown.Length; i++) {
            if (MarkdownSourceColumns.IsLineBreakStart(SourceMarkdown, i, out _)) {
                lineEndExclusive = i;
                break;
            }
        }

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
