namespace OfficeIMO.Markdown;

using System;
using System.Collections.Generic;

/// <summary>
/// Helpers for normalizing blank-line runs in markdown text.
/// </summary>
public static class MarkdownBlankLines {
    /// <summary>
    /// Collapses consecutive blank lines into a single blank line while preserving the input line-ending style.
    /// </summary>
    /// <param name="markdown">Markdown source.</param>
    /// <returns>Markdown with duplicate blank lines collapsed.</returns>
    public static string CollapseDuplicateBlankLines(string? markdown) {
        if (string.IsNullOrEmpty(markdown)) {
            return markdown ?? string.Empty;
        }

        var value = markdown!;
        var newline = DetectLineEnding(value);
        var normalized = value.Replace("\r\n", "\n").Replace("\r", "\n");
        var lines = normalized.Split('\n');
        var output = new List<string>(lines.Length);
        var previousWasBlank = false;

        for (var i = 0; i < lines.Length; i++) {
            var line = lines[i] ?? string.Empty;
            var isBlank = string.IsNullOrWhiteSpace(line);
            if (isBlank && previousWasBlank) {
                continue;
            }

            output.Add(line);
            previousWasBlank = isBlank;
        }

        return string.Join(newline, output);
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
