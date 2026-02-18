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

    /// <summary>
    /// When true, trims accidental whitespace immediately inside strong delimiters
    /// (for example, <c>** Healthy**</c> or <c>**Healthy **</c> become <c>**Healthy**</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeLooseStrongDelimiters { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing space after an ordered list marker when the content starts with
    /// emphasis-like characters (for example, <c>2.**Task**</c> becomes <c>2. **Task**</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeOrderedListMarkerSpacing { get; set; } = false;

    /// <summary>
    /// When true, converts ordered list markers in <c>1)</c> form to <c>1.</c> with normalized spacing.
    /// Default: false.
    /// </summary>
    public bool NormalizeOrderedListParenMarkers { get; set; } = false;

    /// <summary>
    /// When true, removes stray caret artifacts after ordered list markers
    /// (for example, <c>2.^ **Task**</c> becomes <c>2. **Task**</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeOrderedListCaretArtifacts { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing space before parenthetical phrases adjacent to prose or strong spans
    /// (for example, <c>**Task**(detail)</c> becomes <c>**Task** (detail)</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeTightParentheticalSpacing { get; set; } = false;

    /// <summary>
    /// When true, flattens malformed nested strong delimiters emitted by some model outputs
    /// (for example, <c>**from **Service Control Manager**.**</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeNestedStrongDelimiters { get; set; } = false;
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
        @"(\*\*[^\s*\r\n](?:[^*\r\n]*[^\s*\r\n])?\*\*)(?=[\p{L}\p{N}])",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LooseStrongDelimiterWhitespaceRegex = new Regex(
        @"\*\*(?<inner>[^*\r\n]+)\*\*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex OrderedListMarkerMissingSpaceRegex = new Regex(
        @"^(?<prefix>[ \t]{0,3}\d+[.)])(?=[*_`\[])",
        RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline);

    private static readonly Regex OrderedListParenMarkerRegex = new Regex(
        @"^(?<indent>[ \t]{0,3})(?<num>\d+)\)\s*(?=\S)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline);

    private static readonly Regex OrderedListCaretArtifactRegex = new Regex(
        @"^(?<lead>[ \t]{0,3}\d+\.)\s*\^\s*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline);

    private static readonly Regex TightParentheticalSpacingRegex = new Regex(
        @"(?:(?<=\*\*)|(?<=[\p{L}\p{N}\)]))\((?=[\p{L}][^\r\n)]*\))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex NestedStrongSpanRegex = new Regex(
        @"(?<!\S)\*\*(?<left>[^*\r\n]{6,}?\s)\*\*(?<inner>[A-Za-z0-9`][^*:\r\n]*?)\*\*(?<right>[^*\r\n]*?)\*\*",
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

                // Avoid collapsing list boundaries such as:
                // **First item text**
                // 2.** Second item**
                if (LooksLikeOrderedListMarkerFragment(right)) {
                    return match.Value;
                }

                return "**" + left + " " + right + "**";
            });
        }

        if (options.NormalizeNestedStrongDelimiters) {
            value = FlattenNestedStrongSpansOutsideFencedCodeBlocks(value);
        }

        if (options.NormalizeLooseStrongDelimiters) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, LooseStrongDelimiterWhitespaceRegex, static match => {
                var inner = match.Groups["inner"].Value;
                var trimmed = inner.Trim();
                if (trimmed.Length == 0 || trimmed.Length == inner.Length) {
                    return match.Value;
                }

                return "**" + trimmed + "**";
            });
        }

        if (options.NormalizeTightStrongBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, TightStrongSuffixRegex, static match => match.Groups[1].Value + " ");
        }

        if (options.NormalizeOrderedListMarkerSpacing) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, OrderedListMarkerMissingSpaceRegex, static match => match.Groups["prefix"].Value + " ");
        }

        if (options.NormalizeOrderedListParenMarkers) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, OrderedListParenMarkerRegex, static match => match.Groups["indent"].Value + match.Groups["num"].Value + ". ");
        }

        if (options.NormalizeOrderedListCaretArtifacts) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, OrderedListCaretArtifactRegex, static match => match.Groups["lead"].Value + " ");
        }

        if (options.NormalizeTightParentheticalSpacing) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                TightParentheticalSpacingRegex,
                static _ => " (",
                preserveInlineCodeSpans: true);
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

    private static string FlattenNestedStrongSpansOutsideFencedCodeBlocks(string value) {
        var current = value ?? string.Empty;
        while (true) {
            var flattened = ApplyRegexOutsideFencedCodeBlocks(
                current,
                NestedStrongSpanRegex,
                static match =>
                    "**"
                    + match.Groups["left"].Value
                    + match.Groups["inner"].Value
                    + match.Groups["right"].Value
                    + "**");
            if (flattened == current) {
                return flattened;
            }

            current = flattened;
        }
    }

    private static bool LooksLikeOrderedListMarkerFragment(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var trimmed = value.Trim();
        if (trimmed.Length < 2) {
            return false;
        }

        int index = 0;
        while (index < trimmed.Length && char.IsDigit(trimmed[index])) {
            index++;
        }

        if (index == 0 || index != trimmed.Length - 1) {
            return false;
        }

        return trimmed[index] == '.' || trimmed[index] == ')';
    }

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

            if (MarkdownFence.TryReadFenceRun(line, out var runMarker, out var runLength, out var runSuffix)) {
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
