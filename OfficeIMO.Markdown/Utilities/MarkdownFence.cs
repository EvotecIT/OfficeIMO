namespace OfficeIMO.Markdown;

using System;
using System.Text;

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
    /// Tries to parse a fence run from a line that may be prefixed by blockquote container markers
    /// and indentation, returning the preserved prefix separately.
    /// </summary>
    /// <param name="line">Line to inspect.</param>
    /// <param name="linePrefix">Leading indentation and blockquote markers preceding the fence.</param>
    /// <param name="marker">Fence marker (<c>`</c> or <c>~</c>).</param>
    /// <param name="runLength">Fence run length.</param>
    /// <param name="suffix">Any trailing text after the marker run (for example language token).</param>
    /// <returns><c>true</c> when the line contains a valid fence run after container prefixes; otherwise <c>false</c>.</returns>
    public static bool TryReadContainerAwareFenceRun(string? line, out string linePrefix, out char marker, out int runLength, out string suffix) {
        linePrefix = string.Empty;
        marker = '\0';
        runLength = 0;
        suffix = string.Empty;

        if (line is null) {
            return false;
        }

        var fenceCandidateStart = GetContainerAwareFenceCandidateStartIndex(line);
        if (fenceCandidateStart >= line.Length) {
            return false;
        }

        if (!TryReadFenceRun(line.Substring(fenceCandidateStart), out marker, out runLength, out suffix)) {
            return false;
        }

        linePrefix = line.Substring(0, fenceCandidateStart);
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
    /// Applies a text transformation only to segments outside fenced code blocks while preserving
    /// container-aware fence boundaries such as blockquoted or indented fences.
    /// </summary>
    /// <param name="input">Markdown text to transform.</param>
    /// <param name="transformer">Transform to apply outside fenced code blocks.</param>
    /// <returns>Transformed markdown with fenced content preserved verbatim.</returns>
    public static string ApplyTransformOutsideFencedCodeBlocks(string? input, Func<string, string> transformer) {
        if (transformer == null) {
            throw new ArgumentNullException(nameof(transformer));
        }

        if (string.IsNullOrEmpty(input)) {
            return input ?? string.Empty;
        }

        var value = input!;
        var output = new StringBuilder(value.Length);
        var outsideSegment = new StringBuilder();
        var inFence = false;
        var fenceMarker = '\0';
        var fenceRunLength = 0;

        var index = 0;
        while (index < value.Length) {
            var lineStart = index;
            while (index < value.Length && value[index] != '\r' && value[index] != '\n') {
                index++;
            }

            var lineEnd = index;
            if (index < value.Length && value[index] == '\r') {
                index++;
                if (index < value.Length && value[index] == '\n') {
                    index++;
                }
            } else if (index < value.Length && value[index] == '\n') {
                index++;
            }

            var line = value.Substring(lineStart, lineEnd - lineStart);
            var lineWithNewline = value.Substring(lineStart, index - lineStart);

            if (TryReadContainerAwareFenceRun(line, out _, out var runMarker, out var runLength, out var runSuffix)) {
                if (!inFence) {
                    FlushTransformedOutsideSegment(output, outsideSegment, transformer);
                    inFence = true;
                    fenceMarker = runMarker;
                    fenceRunLength = runLength;
                    output.Append(lineWithNewline);
                    continue;
                }

                if (runMarker == fenceMarker && runLength >= fenceRunLength && string.IsNullOrWhiteSpace(runSuffix)) {
                    inFence = false;
                    fenceMarker = '\0';
                    fenceRunLength = 0;
                    output.Append(lineWithNewline);
                    continue;
                }
            }

            if (inFence) {
                output.Append(lineWithNewline);
            } else {
                outsideSegment.Append(lineWithNewline);
            }
        }

        FlushTransformedOutsideSegment(output, outsideSegment, transformer);
        return output.ToString();
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

    private static int GetContainerAwareFenceCandidateStartIndex(string line) {
        var index = 0;

        while (true) {
            var segmentStart = index;
            while (index < line.Length && (line[index] == ' ' || line[index] == '\t')) {
                index++;
            }

            if (index < line.Length && line[index] == '>') {
                index++;
                if (index < line.Length && line[index] == ' ') {
                    index++;
                }

                continue;
            }

            index = segmentStart;
            break;
        }

        while (index < line.Length && (line[index] == ' ' || line[index] == '\t')) {
            index++;
        }

        return index;
    }

    private static void FlushTransformedOutsideSegment(StringBuilder output, StringBuilder outsideSegment, Func<string, string> transformer) {
        if (outsideSegment.Length == 0) {
            return;
        }

        output.Append(transformer(outsideSegment.ToString()));
        outsideSegment.Clear();
    }
}
