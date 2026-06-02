using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownInputNormalizer {
    private static string ApplyRegexOutsideFencedCodeBlocks(
        string input,
        Regex regex,
        MatchEvaluator evaluator,
        bool preserveInlineCodeSpans = false) {
        if (string.IsNullOrEmpty(input)) {
            return input ?? string.Empty;
        }

        var output = new StringBuilder(input.Length);
        var outsideSegment = new StringBuilder();
        var inFence = false;
        char fenceMarker = '\0';
        var fenceRunLength = 0;

        var index = 0;
        while (index < input.Length) {
            var lineStart = index;
            while (index < input.Length && input[index] != '\r' && input[index] != '\n') {
                index++;
            }

            var lineEnd = index;
            if (index < input.Length && input[index] == '\r') {
                index++;
                if (index < input.Length && input[index] == '\n') {
                    index++;
                }
            } else if (index < input.Length && input[index] == '\n') {
                index++;
            }

            var line = input.Substring(lineStart, lineEnd - lineStart);
            var lineWithNewline = input.Substring(lineStart, index - lineStart);

            if (MarkdownFence.TryReadContainerAwareFenceRun(line, out _, out var runMarker, out var runLength, out var runSuffix)) {
                if (!inFence) {
                    FlushOutsideSegment(output, outsideSegment, regex, evaluator, preserveInlineCodeSpans);
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

        FlushOutsideSegment(output, outsideSegment, regex, evaluator, preserveInlineCodeSpans);
        return output.ToString();
    }

    private static string ApplyTransformOutsideFencedCodeBlocks(string input, Func<string, string> transformer) {
        return MarkdownFence.ApplyTransformOutsideFencedCodeBlocks(input, transformer);
    }

    private static void FlushOutsideSegment(
        StringBuilder output,
        StringBuilder outsideSegment,
        Regex regex,
        MatchEvaluator evaluator,
        bool preserveInlineCodeSpans) {
        if (outsideSegment.Length == 0) {
            return;
        }

        var segment = outsideSegment.ToString();
        output.Append(preserveInlineCodeSpans
            ? ReplaceOutsideInlineCodeSpans(segment, regex, evaluator)
            : regex.Replace(segment, evaluator));
        outsideSegment.Clear();
    }

    private static void FlushOutsideSegment(
        StringBuilder output,
        StringBuilder outsideSegment,
        Func<string, string> transformer) {
        if (outsideSegment.Length == 0) {
            return;
        }

        output.Append(transformer(outsideSegment.ToString()));
        outsideSegment.Clear();
    }

    private static string ReplaceOutsideInlineCodeSpans(string value, Regex regex, MatchEvaluator evaluator) {
        if (string.IsNullOrEmpty(value) || value.IndexOf('`') < 0) {
            return regex.Replace(value ?? string.Empty, evaluator);
        }

        var matches = InlineCodeSpanRegex.Matches(value);
        if (matches.Count == 0) {
            return regex.Replace(value, evaluator);
        }

        var output = new StringBuilder(value.Length);
        var cursor = 0;
        for (var i = 0; i < matches.Count; i++) {
            var code = matches[i];
            if (code.Index > cursor) {
                output.Append(regex.Replace(value.Substring(cursor, code.Index - cursor), evaluator));
            }

            output.Append(code.Value);
            cursor = code.Index + code.Length;
        }

        if (cursor < value.Length) {
            output.Append(regex.Replace(value.Substring(cursor), evaluator));
        }

        return output.ToString();
    }
}
