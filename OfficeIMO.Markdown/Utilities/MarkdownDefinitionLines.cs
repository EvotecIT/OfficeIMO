namespace OfficeIMO.Markdown;

using System;
using System.Collections.Generic;

/// <summary>
/// Helpers for shaping adjacent <c>Label: value</c> markdown lines into stable paragraph boundaries.
/// </summary>
public static class MarkdownDefinitionLines {
    /// <summary>
    /// Inserts blank lines between adjacent simple definition-like lines while preserving fenced code blocks verbatim.
    /// </summary>
    /// <param name="markdown">Markdown source.</param>
    /// <returns>Markdown with blank lines inserted between targeted adjacent definition-like lines.</returns>
    public static string SeparateAdjacentDefinitionLikeLinesOutsideFencedCodeBlocks(string? markdown) {
        if (string.IsNullOrEmpty(markdown)) {
            return markdown ?? string.Empty;
        }

        return MarkdownFence.ApplyTransformOutsideFencedCodeBlocks(
            markdown,
            static segment => SeparateAdjacentDefinitionLikeLines(segment));
    }

    private static string SeparateAdjacentDefinitionLikeLines(string markdown) {
        if (string.IsNullOrEmpty(markdown) || !ContainsAdjacentDefinitionLikeLines(markdown)) {
            return markdown ?? string.Empty;
        }

        var newline = DetectLineEnding(markdown);
        var normalized = markdown.Replace("\r\n", "\n").Replace("\r", "\n");
        var lines = normalized.Split('\n');
        var output = new List<string>(lines.Length * 2);

        for (var i = 0; i < lines.Length; i++) {
            var current = lines[i] ?? string.Empty;
            output.Add(current);

            if (i + 1 >= lines.Length) {
                continue;
            }

            if (IsSimpleDefinitionLikeLine(current) && IsSimpleDefinitionLikeLine(lines[i + 1])) {
                output.Add(string.Empty);
            }
        }

        return string.Join(newline, output);
    }

    private static bool ContainsAdjacentDefinitionLikeLines(string markdown) {
        var normalized = markdown.Replace("\r\n", "\n").Replace("\r", "\n");
        var lines = normalized.Split('\n');
        if (lines.Length < 2) {
            return false;
        }

        for (var i = 0; i < lines.Length - 1; i++) {
            if (IsSimpleDefinitionLikeLine(lines[i]) && IsSimpleDefinitionLikeLine(lines[i + 1])) {
                return true;
            }
        }

        return false;
    }

    private static bool IsSimpleDefinitionLikeLine(string? line) {
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        var value = line!;
        if (value.StartsWith("    ", StringComparison.Ordinal) || value[0] == '\t') {
            return false;
        }

        var trimmed = value.Trim();
        if (trimmed.Length == 0
            || trimmed.StartsWith("#", StringComparison.Ordinal)
            || trimmed.StartsWith(">", StringComparison.Ordinal)
            || trimmed.StartsWith("|", StringComparison.Ordinal)
            || trimmed.StartsWith("- ", StringComparison.Ordinal)
            || trimmed.StartsWith("* ", StringComparison.Ordinal)
            || trimmed.StartsWith("+ ", StringComparison.Ordinal)
            || StartsWithOrderedListMarker(trimmed)) {
            return false;
        }

        var separatorIndex = FindDefinitionSeparatorIndex(trimmed);
        if (separatorIndex <= 0 || separatorIndex >= trimmed.Length - 1) {
            return false;
        }

        var label = trimmed.Substring(0, separatorIndex).Trim();
        var definitionValue = trimmed.Substring(separatorIndex + 1).Trim();
        if (label.Length == 0 || definitionValue.Length == 0 || label.Length > 48) {
            return false;
        }

        for (var i = 0; i < label.Length; i++) {
            var ch = label[i];
            if (char.IsLetterOrDigit(ch) || ch == ' ' || ch == '-' || ch == '_' || ch == '/' || ch == '`' || ch == '\'') {
                continue;
            }

            return false;
        }

        return true;
    }

    private static int FindDefinitionSeparatorIndex(string line) {
        var inInlineCode = false;
        for (var i = 0; i < line.Length - 1; i++) {
            if (line[i] == '`') {
                inInlineCode = !inInlineCode;
                continue;
            }

            if (!inInlineCode && line[i] == ':' && char.IsWhiteSpace(line[i + 1])) {
                return i;
            }
        }

        return -1;
    }

    private static bool StartsWithOrderedListMarker(string trimmed) {
        var index = 0;
        while (index < trimmed.Length && char.IsDigit(trimmed[index])) {
            index++;
        }

        return index > 0
               && index + 1 < trimmed.Length
               && (trimmed[index] == '.' || trimmed[index] == ')')
               && char.IsWhiteSpace(trimmed[index + 1]);
    }

    private static string DetectLineEnding(string text) {
        if (text.Contains("\r\n")) {
            return "\r\n";
        }

        if (text.IndexOf('\r') >= 0) {
            return "\r";
        }

        return "\n";
    }
}
