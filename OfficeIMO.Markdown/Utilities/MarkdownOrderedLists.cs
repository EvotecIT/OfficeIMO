namespace OfficeIMO.Markdown;

using System;
using System.Text;

/// <summary>
/// Helpers for repairing malformed ordered-list spacing in markdown text.
/// </summary>
public static class MarkdownOrderedLists {
    /// <summary>
    /// Inserts a blank line between adjacent ordered-list items that were emitted without paragraph separation.
    /// </summary>
    /// <param name="markdown">Markdown source.</param>
    /// <returns>Markdown with adjacent ordered-list items separated by a blank line.</returns>
    public static string SeparateAdjacentOrderedListItems(string? markdown) {
        if (string.IsNullOrEmpty(markdown) || markdown!.IndexOf('\n') < 0) {
            return markdown ?? string.Empty;
        }

        return MarkdownFence.ApplyTransformOutsideFencedCodeBlocks(
            markdown,
            static segment => SeparateAdjacentOrderedListItemsCore(segment));
    }

    private static string SeparateAdjacentOrderedListItemsCore(string markdown) {
        var normalized = markdown.Replace("\r\n", "\n").Replace("\r", "\n");
        var lines = normalized.Split('\n');
        if (lines.Length < 2) {
            return normalized;
        }

        var sb = new StringBuilder(normalized.Length + 32);
        for (var i = 0; i < lines.Length; i++) {
            var current = lines[i] ?? string.Empty;
            sb.Append(current);
            if (i >= lines.Length - 1) {
                continue;
            }

            sb.Append('\n');
            var next = lines[i + 1] ?? string.Empty;
            if (IsOrderedListLine(current) && IsOrderedListLine(next)) {
                sb.Append('\n');
            }
        }

        return sb.ToString();
    }

    private static bool IsOrderedListLine(string? line) {
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        var value = line!;
        var i = 0;
        while (i < value.Length && char.IsWhiteSpace(value[i])) {
            i++;
        }

        var numberStart = i;
        while (i < value.Length && char.IsDigit(value[i])) {
            i++;
        }

        if (i == numberStart || i >= value.Length || value[i] != '.') {
            return false;
        }

        i++;
        return i < value.Length && char.IsWhiteSpace(value[i]);
    }
}
