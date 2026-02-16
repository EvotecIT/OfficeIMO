using System.Text.RegularExpressions;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Options for lightweight markdown text normalization before parsing.
/// </summary>
public sealed class MarkdownInputNormalizationOptions {
    /// <summary>
    /// When true, joins short hard-wrapped bold labels (for example, "**Status\nOK**") into a single bold span.
    /// Default: false.
    /// </summary>
    public bool NormalizeSoftWrappedStrongSpans { get; set; } = false;

    /// <summary>
    /// When true, compacts inline code spans containing line breaks into a single line.
    /// Default: false.
    /// </summary>
    public bool NormalizeInlineCodeSpanLineBreaks { get; set; } = false;

    /// <summary>
    /// When true, converts escaped inline code spans (for example, <c>\`code\`</c>) into standard markdown code spans.
    /// This helps chat/model outputs that over-escape backticks.
    /// Default: false.
    /// </summary>
    public bool NormalizeEscapedInlineCodeSpans { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing space after a closing strong span when followed by a word character
    /// (for example, <c>**Healthy**next</c> becomes <c>**Healthy** next</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeTightStrongBoundaries { get; set; } = false;
}

/// <summary>
/// Stateless markdown input normalizer intended for chat/model outputs before strict parsing.
/// </summary>
public static class MarkdownInputNormalizer {
    private static readonly Regex InlineCodeSpanRegex = new Regex(
        "`([^`]+)`",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex SoftWrappedStrongRegex = new Regex(
        "\\*\\*(?<left>[^\\r\\n*]{1,80})\\r?\\n(?<right>[^\\r\\n*]{1,80})\\*\\*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex EscapedInlineCodeSpanRegex = new Regex(
        @"\\`(?<code>[^`\r\n]+?)\\`",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex TightStrongSuffixRegex = new Regex(
        @"(\*\*[^*\r\n]+\*\*)(?=[\p{L}\p{N}])",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// Normalizes markdown text based on <paramref name="options"/>.
    /// </summary>
    /// <param name="markdown">Input markdown.</param>
    /// <param name="options">Normalization options.</param>
    /// <returns>Normalized markdown.</returns>
    public static string Normalize(string? markdown, MarkdownInputNormalizationOptions? options = null) {
        var value = markdown ?? string.Empty;
        if (value.Length == 0) {
            return value;
        }

        options ??= new MarkdownInputNormalizationOptions();

        if (options.NormalizeSoftWrappedStrongSpans) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, SoftWrappedStrongRegex, static match => {
                var left = match.Groups["left"].Value.Trim();
                var right = match.Groups["right"].Value.Trim();
                if (left.Length == 0 || right.Length == 0) {
                    return match.Value;
                }

                return "**" + left + " " + right + "**";
            });
        }

        if (options.NormalizeTightStrongBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, TightStrongSuffixRegex, static match => match.Groups[1].Value + " ");
        }

        if (options.NormalizeInlineCodeSpanLineBreaks) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, InlineCodeSpanRegex, static match => {
                var body = match.Groups[1].Value;
                if (body.IndexOfAny(new[] { '\r', '\n' }) < 0) {
                    return match.Value;
                }

                var compact = body.Replace("\r\n", " ")
                    .Replace('\r', ' ')
                    .Replace('\n', ' ')
                    .Trim();
                return compact.Length == 0 ? "``" : "`" + compact + "`";
            });
        }

        if (options.NormalizeEscapedInlineCodeSpans) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, EscapedInlineCodeSpanRegex, static match => {
                var body = match.Groups["code"].Value;
                return body.Length == 0 ? "``" : "`" + body + "`";
            });
        }

        return value;
    }

    private static string ApplyRegexOutsideFencedCodeBlocks(string input, Regex regex, MatchEvaluator evaluator) {
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

            if (MarkdownFence.TryReadFenceRun(line, out var runMarker, out var runLength, out var runSuffix)) {
                if (!inFence) {
                    FlushOutsideSegment(output, outsideSegment, regex, evaluator);
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

        FlushOutsideSegment(output, outsideSegment, regex, evaluator);
        return output.ToString();
    }

    private static void FlushOutsideSegment(StringBuilder output, StringBuilder outsideSegment, Regex regex, MatchEvaluator evaluator) {
        if (outsideSegment.Length == 0) {
            return;
        }

        output.Append(regex.Replace(outsideSegment.ToString(), evaluator));
        outsideSegment.Clear();
    }
}
