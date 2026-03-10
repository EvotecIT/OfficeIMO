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
    /// When true, normalizes compact arrow-to-strong boundaries
    /// (for example, <c>->**Why it matters:**</c> becomes <c>-> **Why it matters:**</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeTightArrowStrongBoundaries { get; set; } = false;

    /// <summary>
    /// When true, repairs malformed strong spans that are missing the closing delimiter
    /// immediately before an arrow-led strong label
    /// (for example, <c>**No current failures -&gt; **Why it matters:**</c> becomes
    /// <c>**No current failures** -&gt; **Why it matters:**</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeBrokenStrongArrowLabels { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing space after a colon in prose labels
    /// (for example, <c>Why it matters:missing coverage</c> becomes <c>Why it matters: missing coverage</c>).
    /// This is applied by AST-level inline normalization and intentionally skips inline code spans.
    /// Default: false.
    /// </summary>
    public bool NormalizeTightColonSpacing { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing newline between an ATX heading and an immediately-following
    /// unordered strong-label list marker on the same line
    /// (for example, <c>## Summary- **Item:** value</c> becomes <c>## Summary\n- **Item:** value</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeHeadingListBoundaries { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing newline before compact unordered strong-label list markers
    /// that were emitted inline after punctuation or symbol characters
    /// (for example, <c>✅- **FSMO:** ok</c> becomes <c>✅\n- **FSMO:** ok</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeCompactStrongLabelListBoundaries { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing newline before compact ATX headings emitted directly after prose or symbols
    /// on the same line (for example, <c>unexpected### Reason</c> becomes <c>unexpected\n### Reason</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeCompactHeadingBoundaries { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing newline between a colon and an immediately-following unordered list marker
    /// on the same line (for example, <c>Next step:- **Item**</c> becomes <c>Next step:\n- **Item**</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeColonListBoundaries { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing newline between a fenced code block language token and inline body content
    /// for common compact model-output mistakes
    /// (for example, <c>```json{"x":1}</c> becomes <c>```json\n{"x":1}</c>,
    /// and <c>```mermaidflowchart LR A--&gt;B</c> becomes <c>```mermaid\nflowchart LR A--&gt;B</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeCompactFenceBodyBoundaries { get; set; } = false;

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

    /// <summary>
    /// Enables the named preset on the current options instance.
    /// Existing values are overwritten by the preset defaults.
    /// </summary>
    /// <param name="preset">Preset to apply.</param>
    /// <returns>The same options instance for chaining.</returns>
    public MarkdownInputNormalizationOptions ApplyPreset(MarkdownInputNormalizationPreset preset) {
        MarkdownInputNormalizationPresets.ApplyTo(this, preset);
        return this;
    }
}

/// <summary>
/// Named input-normalization presets for common markdown ingestion scenarios.
/// </summary>
public enum MarkdownInputNormalizationPreset {
    /// <summary>No preset behavior.</summary>
    None = 0,
    /// <summary>
    /// Conservative transcript repair preset aligned with current chat-host bridge behavior.
    /// </summary>
    ChatTranscript = 1,
    /// <summary>
    /// Broader chat/model-output repair preset for aggressively malformed transcript content.
    /// </summary>
    ChatStrict = 2,
    /// <summary>
    /// Conservative documentation import preset that avoids transcript-specific boundary rewrites.
    /// </summary>
    DocsLoose = 3
}

/// <summary>
/// Factory and application helpers for <see cref="MarkdownInputNormalizationPreset"/>.
/// </summary>
public static class MarkdownInputNormalizationPresets {
    /// <summary>
    /// Creates a new options instance with the selected preset applied.
    /// </summary>
    /// <param name="preset">Preset to create.</param>
    /// <returns>New options instance.</returns>
    public static MarkdownInputNormalizationOptions Create(MarkdownInputNormalizationPreset preset) {
        var options = new MarkdownInputNormalizationOptions();
        ApplyTo(options, preset);
        return options;
    }

    /// <summary>
    /// Creates the conservative transcript-repair preset aligned with existing chat bridge behavior.
    /// </summary>
    public static MarkdownInputNormalizationOptions CreateChatTranscript() {
        return Create(MarkdownInputNormalizationPreset.ChatTranscript);
    }

    /// <summary>
    /// Creates the broader chat/model-output repair preset.
    /// </summary>
    public static MarkdownInputNormalizationOptions CreateChatStrict() {
        return Create(MarkdownInputNormalizationPreset.ChatStrict);
    }

    /// <summary>
    /// Creates the conservative documentation import preset.
    /// </summary>
    public static MarkdownInputNormalizationOptions CreateDocsLoose() {
        return Create(MarkdownInputNormalizationPreset.DocsLoose);
    }

    /// <summary>
    /// Applies a preset to an existing options instance.
    /// Existing values are overwritten by the preset defaults.
    /// </summary>
    /// <param name="options">Options instance to mutate.</param>
    /// <param name="preset">Preset to apply.</param>
    public static void ApplyTo(MarkdownInputNormalizationOptions options, MarkdownInputNormalizationPreset preset) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        Reset(options);

        switch (preset) {
            case MarkdownInputNormalizationPreset.None:
                return;
            case MarkdownInputNormalizationPreset.ChatTranscript:
                ApplyChatTranscript(options);
                return;
            case MarkdownInputNormalizationPreset.ChatStrict:
                ApplyChatStrict(options);
                return;
            case MarkdownInputNormalizationPreset.DocsLoose:
                ApplyDocsLoose(options);
                return;
            default:
                throw new ArgumentOutOfRangeException(nameof(preset), preset, "Unknown markdown input normalization preset.");
        }
    }

    private static void Reset(MarkdownInputNormalizationOptions options) {
        options.NormalizeSoftWrappedStrongSpans = false;
        options.NormalizeInlineCodeSpanLineBreaks = false;
        options.NormalizeEscapedInlineCodeSpans = false;
        options.NormalizeTightStrongBoundaries = false;
        options.NormalizeTightArrowStrongBoundaries = false;
        options.NormalizeBrokenStrongArrowLabels = false;
        options.NormalizeTightColonSpacing = false;
        options.NormalizeHeadingListBoundaries = false;
        options.NormalizeCompactStrongLabelListBoundaries = false;
        options.NormalizeCompactHeadingBoundaries = false;
        options.NormalizeColonListBoundaries = false;
        options.NormalizeCompactFenceBodyBoundaries = false;
        options.NormalizeLooseStrongDelimiters = false;
        options.NormalizeOrderedListMarkerSpacing = false;
        options.NormalizeOrderedListParenMarkers = false;
        options.NormalizeOrderedListCaretArtifacts = false;
        options.NormalizeTightParentheticalSpacing = false;
        options.NormalizeNestedStrongDelimiters = false;
    }

    private static void ApplyChatTranscript(MarkdownInputNormalizationOptions options) {
        options.NormalizeLooseStrongDelimiters = true;
        options.NormalizeTightStrongBoundaries = true;
        options.NormalizeOrderedListMarkerSpacing = true;
        options.NormalizeOrderedListParenMarkers = true;
        options.NormalizeOrderedListCaretArtifacts = true;
        options.NormalizeTightParentheticalSpacing = true;
        options.NormalizeNestedStrongDelimiters = true;
        options.NormalizeTightArrowStrongBoundaries = true;
        options.NormalizeTightColonSpacing = true;
    }

    private static void ApplyChatStrict(MarkdownInputNormalizationOptions options) {
        ApplyChatTranscript(options);
        options.NormalizeSoftWrappedStrongSpans = true;
        options.NormalizeInlineCodeSpanLineBreaks = true;
        options.NormalizeEscapedInlineCodeSpans = true;
        options.NormalizeBrokenStrongArrowLabels = true;
        options.NormalizeHeadingListBoundaries = true;
        options.NormalizeCompactStrongLabelListBoundaries = true;
        options.NormalizeCompactHeadingBoundaries = true;
        options.NormalizeColonListBoundaries = true;
        options.NormalizeCompactFenceBodyBoundaries = true;
    }

    private static void ApplyDocsLoose(MarkdownInputNormalizationOptions options) {
        options.NormalizeLooseStrongDelimiters = true;
        options.NormalizeTightStrongBoundaries = true;
        options.NormalizeOrderedListMarkerSpacing = true;
        options.NormalizeOrderedListParenMarkers = true;
        options.NormalizeOrderedListCaretArtifacts = true;
        options.NormalizeTightParentheticalSpacing = true;
        options.NormalizeNestedStrongDelimiters = true;
    }
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

    private static readonly Regex TightArrowStrongBoundaryRegex = new Regex(
        @"->\s*(?=\*\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex BrokenStrongArrowLabelRegex = new Regex(
        @"\*\*(?<left>[^*\r\n]{1,200}?)\s*->\s*\*\*(?<label>[^*\r\n]{1,120}?):\*\*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex HeadingListBoundaryRegex = new Regex(
        @"^(?<heading>[ \t]{0,3}#{1,6}[ \t]+[^\r\n]+?)(?<!\s)(?<marker>[-+*])\s+(?=\*\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline);

    private static readonly Regex CompactStrongLabelListBoundaryRegex = new Regex(
        @"(?<=[\p{P}\p{S}\)])(?<marker>[-+*])\s+(?=\*\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex CompactHeadingBoundaryRegex = new Regex(
        @"(?<=[^\s\r\n])(?<marker>#{2,6})\s+(?=\S)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex ColonListBoundaryRegex = new Regex(
        @":\s*(?<marker>[-+*])\s+(?=(\*\*|`|\[|\p{L}|\p{N}))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly string[] CompactFenceLanguages = {
        "ix-dataview",
        "ix-network",
        "visnetwork",
        "ix-chart",
        "mermaid",
        "network",
        "chart",
        "jsonc",
        "json5",
        "json"
    };

    private static readonly string[] MermaidBodyPrefixes = {
        "flowchart",
        "graph",
        "sequencediagram",
        "classdiagram",
        "statediagram-v2",
        "statediagram",
        "erdiagram",
        "journey",
        "gantt",
        "pie",
        "mindmap",
        "timeline",
        "quadrantchart",
        "xychart",
        "sankey-beta",
        "requirement",
        "gitgraph",
        "c4context",
        "c4container",
        "c4component",
        "c4dynamic",
        "c4deployment"
    };

    private static readonly Regex LooseStrongDelimiterWhitespaceRegex = new Regex(
        @"\*\*(?<inner>[^*\r\n]+)\*\*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex RepeatedStrongDelimiterRunRegex = new Regex(
        @"(?<!\*)(?<left>\*{4,})(?<inner>[^*\r\n]+?)(?<right>\*{4,})(?!\*)",
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
                if (LooksLikeListMarkerFragment(right)) {
                    return match.Value;
                }

                return "**" + left + " " + right + "**";
            });
        }

        if (options.NormalizeNestedStrongDelimiters) {
            value = FlattenNestedStrongSpansOutsideFencedCodeBlocks(value);
        }

        if (options.NormalizeLooseStrongDelimiters) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, RepeatedStrongDelimiterRunRegex, static match => {
                var leftLength = match.Groups["left"].Value.Length;
                var rightLength = match.Groups["right"].Value.Length;
                if (leftLength != rightLength || leftLength % 2 != 0) {
                    return match.Value;
                }

                var inner = match.Groups["inner"].Value;
                var trimmed = inner.Trim();
                if (trimmed.Length == 0) {
                    return match.Value;
                }

                return "**" + trimmed + "**";
            });

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

        if (options.NormalizeTightArrowStrongBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                TightArrowStrongBoundaryRegex,
                static _ => "-> ",
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeBrokenStrongArrowLabels) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                BrokenStrongArrowLabelRegex,
                static match => {
                    var left = match.Groups["left"].Value.Trim();
                    var label = match.Groups["label"].Value.Trim();
                    if (left.Length == 0 || label.Length == 0) {
                        return match.Value;
                    }

                    return "**" + left + "** -> **" + label + ":**";
                },
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeCompactHeadingBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                CompactHeadingBoundaryRegex,
                static match => "\n" + match.Groups["marker"].Value + " ",
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeHeadingListBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                HeadingListBoundaryRegex,
                static match => match.Groups["heading"].Value.TrimEnd() + "\n" + match.Groups["marker"].Value + " ",
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeCompactStrongLabelListBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                CompactStrongLabelListBoundaryRegex,
                static match => "\n" + match.Groups["marker"].Value + " ",
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeColonListBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                ColonListBoundaryRegex,
                static match => ":\n" + match.Groups["marker"].Value + " ",
                preserveInlineCodeSpans: true);
        }

        if (options.NormalizeCompactFenceBodyBoundaries) {
            value = NormalizeCompactFenceBodyBoundaries(value);
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

    private static bool LooksLikeListMarkerFragment(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var leading = value.TrimStart();
        if (leading.Length == 0) {
            return false;
        }

        // Unordered list boundaries captured as right-side fragments after a newline.
        // Example captured values: "- ", "-  ", "-".
        if (leading[0] == '-' || leading[0] == '*' || leading[0] == '+') {
            if (leading.Length == 1) {
                return true;
            }

            var next = leading[1];
            if (char.IsWhiteSpace(next) || next == '*' || next == '`') {
                return true;
            }
        }

        var trimmed = leading.Trim();
        if (trimmed.Length < 2) {
            return false;
        }

        // Ordered list boundaries (for example, "2.** ...", "2) ...").
        int index = 0;
        while (index < trimmed.Length && char.IsDigit(trimmed[index])) {
            index++;
        }

        if (index == 0 || index != trimmed.Length - 1) {
            return false;
        }

        return trimmed[index] == '.' || trimmed[index] == ')';
    }

    private static string NormalizeCompactFenceBodyBoundaries(string input) {
        if (string.IsNullOrEmpty(input)) {
            return input ?? string.Empty;
        }

        var output = new StringBuilder(input.Length + 64);
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
            var newline = input.Substring(lineEnd, index - lineEnd);

            if (!inFence && TryNormalizeCompactFenceOpeningLine(line, out var normalizedLine, out var normalizedMarker, out var normalizedRunLength)) {
                inFence = true;
                fenceMarker = normalizedMarker;
                fenceRunLength = normalizedRunLength;
                output.Append(normalizedLine);
                output.Append(newline);
                continue;
            }

            if (MarkdownFence.TryReadFenceRun(line, out var runMarker, out var runLength, out var runSuffix)) {
                if (!inFence) {
                    inFence = true;
                    fenceMarker = runMarker;
                    fenceRunLength = runLength;
                } else if (runMarker == fenceMarker && runLength >= fenceRunLength && string.IsNullOrWhiteSpace(runSuffix)) {
                    inFence = false;
                    fenceMarker = '\0';
                    fenceRunLength = 0;
                }
            }

            output.Append(line);
            output.Append(newline);
        }

        return output.ToString();
    }

    private static bool TryNormalizeCompactFenceOpeningLine(string line, out string normalizedLine, out char fenceMarker, out int fenceRunLength) {
        normalizedLine = line ?? string.Empty;
        fenceMarker = '\0';
        fenceRunLength = 0;

        if (string.IsNullOrEmpty(line)) {
            return false;
        }

        if (!MarkdownFence.TryReadFenceRun(line, out fenceMarker, out fenceRunLength, out var runSuffix)) {
            return false;
        }

        if (string.IsNullOrWhiteSpace(runSuffix)) {
            return false;
        }

        var suffix = runSuffix.TrimStart();
        if (suffix.Length == 0) {
            return false;
        }

        if (!TrySplitCompactFenceSuffix(suffix, out var language, out var body)) {
            return false;
        }

        normalizedLine = new string(fenceMarker, fenceRunLength) + language + "\n" + body;
        return true;
    }

    private static bool TrySplitCompactFenceSuffix(string suffix, out string language, out string body) {
        language = string.Empty;
        body = string.Empty;

        foreach (var candidate in CompactFenceLanguages) {
            if (!suffix.StartsWith(candidate, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            var remainder = suffix.Substring(candidate.Length);
            if (string.IsNullOrWhiteSpace(remainder)) {
                continue;
            }

            if (!LooksLikeCompactFenceBody(candidate, remainder)) {
                continue;
            }

            language = candidate;
            body = remainder;
            return true;
        }

        return false;
    }

    private static bool LooksLikeCompactFenceBody(string language, string remainder) {
        if (string.IsNullOrWhiteSpace(remainder)) {
            return false;
        }

        var trimmed = remainder.TrimStart();
        if (trimmed.Length == 0) {
            return false;
        }

        if (string.Equals(language, "mermaid", StringComparison.OrdinalIgnoreCase)) {
            foreach (var prefix in MermaidBodyPrefixes) {
                if (trimmed.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }

            return false;
        }

        if (string.Equals(language, "json", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "jsonc", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "json5", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "chart", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "ix-chart", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "network", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "visnetwork", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "ix-network", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(language, "ix-dataview", StringComparison.OrdinalIgnoreCase)) {
            return trimmed[0] == '{' || trimmed[0] == '[';
        }

        return false;
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
