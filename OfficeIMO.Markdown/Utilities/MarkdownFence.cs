namespace OfficeIMO.Markdown;

using System;

/// <summary>
/// Helpers for reading markdown fenced-code boundaries and choosing non-colliding fences.
/// </summary>
public static class MarkdownFence {
    /// <summary>
    /// Tries to parse a fence run from a markdown line.
    /// </summary>
    /// <param name="line">Line to inspect.</param>
    /// <param name="marker">Fence marker (<c>`</c> or <c>~</c>).</param>
    /// <param name="runLength">Fence run length.</param>
    /// <param name="suffix">Any trailing text after the marker run (for example language token).</param>
    /// <returns><c>true</c> when the line starts with a valid fence run; otherwise <c>false</c>.</returns>
    public static bool TryReadFenceRun(string? line, out char marker, out int runLength, out string suffix) {
        marker = '\0';
        runLength = 0;
        suffix = string.Empty;
        if (line is null) {
            return false;
        }

        var trimmed = line.TrimStart();
        if (trimmed.Length < 3) {
            return false;
        }

        var first = trimmed[0];
        if (first != '`' && first != '~') {
            return false;
        }

        var i = 0;
        while (i < trimmed.Length && trimmed[i] == first) {
            i++;
        }

        if (i < 3) {
            return false;
        }

        marker = first;
        runLength = i;
        suffix = trimmed.Substring(i);
        return true;
    }

    /// <summary>
    /// Builds a code-fence marker that cannot be prematurely closed by runs inside <paramref name="content"/>.
    /// </summary>
    /// <param name="content">Code content.</param>
    /// <param name="minimumLength">Minimum fence length (defaults to 3).</param>
    /// <returns>Safe fence marker (backticks or tildes).</returns>
    public static string BuildSafeFence(string? content, int minimumLength = 3) {
        var value = content ?? string.Empty;
        var longestBackticks = LongestRun(value, '`');
        var longestTildes = LongestRun(value, '~');

        var backtickLength = Math.Max(minimumLength, longestBackticks + 1);
        var tildeLength = Math.Max(minimumLength, longestTildes + 1);
        var marker = backtickLength <= tildeLength ? '`' : '~';
        var length = marker == '`' ? backtickLength : tildeLength;
        return new string(marker, length);
    }

    /// <summary>
    /// Returns the longest contiguous run of <paramref name="marker"/> within <paramref name="text"/>.
    /// </summary>
    public static int LongestRun(string? text, char marker) {
        if (text == null || text.Length == 0) {
            return 0;
        }

        var longest = 0;
        var current = 0;
        for (var i = 0; i < text.Length; i++) {
            if (text[i] == marker) {
                current++;
                if (current > longest) {
                    longest = current;
                }
            } else {
                current = 0;
            }
        }

        return longest;
    }
}
