using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Options for lightweight markdown text normalization before parsing.
/// </summary>
public sealed class MarkdownInputNormalizationOptions {
    /// <summary>
    /// When true, removes zero-width spacing artifacts such as U+200B/U+2060/U+FEFF that can break markdown readability.
    /// Default: false.
    /// </summary>
    public bool NormalizeZeroWidthSpacingArtifacts { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing space between common status/emoji markers and following prose
    /// (for example, <c>✅Healthy</c> becomes <c>✅ Healthy</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeEmojiWordJoins { get; set; } = false;

    /// <summary>
    /// When true, inserts missing spacing around compact numbered-choice joins
    /// (for example, <c>or2)</c> becomes <c>or 2)</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeCompactNumberedChoiceBoundaries { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing newline before compact bullet markers emitted directly after sentence punctuation
    /// (for example, <c>Done.- **Next:** check</c> becomes <c>Done.\n- **Next:** check</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeSentenceCollapsedBullets { get; set; } = false;

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
    /// When true, repairs malformed signal-flow bullets where an entire arrow chain was accidentally wrapped
    /// in one strong span.
    /// Default: false.
    /// </summary>
    public bool NormalizeWrappedSignalFlowStrongRuns { get; set; } = false;

    /// <summary>
    /// When true, inserts missing spaces after signal-flow labels inside arrow segments
    /// (for example, <c>-> **Why it matters:**coverage</c> becomes
    /// <c>-> **Why it matters:** coverage</c>, and <c>-> Why it matters:coverage</c> becomes
    /// <c>-> Why it matters: coverage</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeSignalFlowLabelSpacing { get; set; } = false;

    /// <summary>
    /// When true, expands collapsed transcript-style metric chains into real markdown lines and
    /// converts legacy bold metric labels into plain labels with bold values.
    /// Default: false.
    /// </summary>
    public bool NormalizeCollapsedMetricChains { get; set; } = false;

    /// <summary>
    /// When true, repairs compact host-label bullets and merges plain continuation lines
    /// (for example, <c>-AD1</c> followed by <c>healthy</c> becomes <c>- AD1 healthy</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeHostLabelBulletArtifacts { get; set; } = false;

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
    /// <see cref="MarkdownReader"/> prefers to apply this via a document transform after parse when the markdown
    /// already parsed into a recoverable paragraph block.
    /// Default: false.
    /// </summary>
    public bool NormalizeCompactHeadingBoundaries { get; set; } = false;

    /// <summary>
    /// When true, removes stray standalone <c>#</c> separator lines that appear immediately before a real ATX heading.
    /// <see cref="MarkdownReader"/> prefers to apply this via a document transform after parse so hosts can
    /// keep the repair in the AST pipeline when the markdown already parses cleanly.
    /// Default: false.
    /// </summary>
    public bool NormalizeStandaloneHashHeadingSeparators { get; set; } = false;

    /// <summary>
    /// When true, repairs broken two-line strong lead-ins such as
    /// <c>**Result</c> followed by <c>body text** tail</c>.
    /// Default: false.
    /// </summary>
    public bool NormalizeBrokenTwoLineStrongLeadIns { get; set; } = false;

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
    /// When true, inserts missing newlines between adjacent ordered list items that were emitted on one line
    /// (for example, <c>...)</c><c>2.**Task**</c> becomes <c>...)\n2.**Task**</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeCollapsedOrderedListBoundaries { get; set; } = false;

    /// <summary>
    /// When true, repairs malformed ordered list items where a bold title is missing its closing strong delimiter
    /// before a parenthetical detail section
    /// (for example, <c>1. **Task(detail)</c> becomes <c>1. **Task** (detail)</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeOrderedListStrongDetailClosures { get; set; } = false;

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
    /// When true, upgrades trailing list-item strong-close artifacts
    /// (for example, <c>- Overall health Healthy****</c> becomes <c>- Overall health **Healthy**</c>).
    /// Default: false.
    /// </summary>
    public bool NormalizeDanglingTrailingStrongListClosers { get; set; } = false;

    /// <summary>
    /// When true, repairs malformed strong runs used as metric values
    /// (for example, <c>******healthy**</c>, <c>**✅****Healthy**</c>, or a missing trailing closer).
    /// Default: false.
    /// </summary>
    public bool NormalizeMetricValueStrongRuns { get; set; } = false;

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
     /// Conservative transcript repair preset aligned with the explicit IntelligenceX transcript contract.
     /// </summary>
    IntelligenceXTranscript = 1,
    /// <summary>
    /// Broader IntelligenceX transcript repair preset for aggressively malformed transcript content.
    /// </summary>
    IntelligenceXTranscriptStrict = 2,
    /// <summary>
     /// Conservative documentation import preset that avoids transcript-specific boundary rewrites.
     /// </summary>
    DocsLoose = 3,
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
    /// Creates the explicit IntelligenceX transcript contract preset.
    /// </summary>
    public static MarkdownInputNormalizationOptions CreateIntelligenceXTranscript() {
        return Create(MarkdownInputNormalizationPreset.IntelligenceXTranscript);
    }

    /// <summary>
    /// Creates the broader IntelligenceX transcript repair preset.
    /// </summary>
    public static MarkdownInputNormalizationOptions CreateIntelligenceXTranscriptStrict() {
        return Create(MarkdownInputNormalizationPreset.IntelligenceXTranscriptStrict);
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
            case MarkdownInputNormalizationPreset.DocsLoose:
                ApplyDocsLoose(options);
                return;
            case MarkdownInputNormalizationPreset.IntelligenceXTranscript:
                ApplyIntelligenceXTranscript(options);
                return;
            case MarkdownInputNormalizationPreset.IntelligenceXTranscriptStrict:
                ApplyIntelligenceXTranscriptStrict(options);
                return;
            default:
                throw new ArgumentOutOfRangeException(nameof(preset), preset, "Unknown markdown input normalization preset.");
        }
    }

    private static void Reset(MarkdownInputNormalizationOptions options) {
        options.NormalizeZeroWidthSpacingArtifacts = false;
        options.NormalizeEmojiWordJoins = false;
        options.NormalizeCompactNumberedChoiceBoundaries = false;
        options.NormalizeSentenceCollapsedBullets = false;
        options.NormalizeSoftWrappedStrongSpans = false;
        options.NormalizeInlineCodeSpanLineBreaks = false;
        options.NormalizeEscapedInlineCodeSpans = false;
        options.NormalizeTightStrongBoundaries = false;
        options.NormalizeTightArrowStrongBoundaries = false;
        options.NormalizeBrokenStrongArrowLabels = false;
        options.NormalizeWrappedSignalFlowStrongRuns = false;
        options.NormalizeSignalFlowLabelSpacing = false;
        options.NormalizeCollapsedMetricChains = false;
        options.NormalizeHostLabelBulletArtifacts = false;
        options.NormalizeTightColonSpacing = false;
        options.NormalizeHeadingListBoundaries = false;
        options.NormalizeCompactStrongLabelListBoundaries = false;
        options.NormalizeCompactHeadingBoundaries = false;
        options.NormalizeStandaloneHashHeadingSeparators = false;
        options.NormalizeBrokenTwoLineStrongLeadIns = false;
        options.NormalizeColonListBoundaries = false;
        options.NormalizeCompactFenceBodyBoundaries = false;
        options.NormalizeLooseStrongDelimiters = false;
        options.NormalizeOrderedListMarkerSpacing = false;
        options.NormalizeOrderedListParenMarkers = false;
        options.NormalizeOrderedListCaretArtifacts = false;
        options.NormalizeCollapsedOrderedListBoundaries = false;
        options.NormalizeOrderedListStrongDetailClosures = false;
        options.NormalizeTightParentheticalSpacing = false;
        options.NormalizeNestedStrongDelimiters = false;
        options.NormalizeDanglingTrailingStrongListClosers = false;
        options.NormalizeMetricValueStrongRuns = false;
    }

    private static void ApplyIntelligenceXTranscript(MarkdownInputNormalizationOptions options) {
        options.NormalizeZeroWidthSpacingArtifacts = true;
        options.NormalizeEmojiWordJoins = true;
        options.NormalizeCompactNumberedChoiceBoundaries = true;
        options.NormalizeSentenceCollapsedBullets = true;
        options.NormalizeLooseStrongDelimiters = true;
        options.NormalizeTightStrongBoundaries = true;
        options.NormalizeOrderedListMarkerSpacing = true;
        options.NormalizeOrderedListParenMarkers = true;
        options.NormalizeOrderedListCaretArtifacts = true;
        options.NormalizeCollapsedOrderedListBoundaries = true;
        options.NormalizeOrderedListStrongDetailClosures = true;
        options.NormalizeTightParentheticalSpacing = true;
        options.NormalizeNestedStrongDelimiters = true;
        options.NormalizeTightArrowStrongBoundaries = true;
        options.NormalizeTightColonSpacing = true;
        options.NormalizeWrappedSignalFlowStrongRuns = true;
        options.NormalizeSignalFlowLabelSpacing = true;
        options.NormalizeCollapsedMetricChains = true;
        options.NormalizeHostLabelBulletArtifacts = true;
        options.NormalizeStandaloneHashHeadingSeparators = true;
        options.NormalizeBrokenTwoLineStrongLeadIns = true;
        options.NormalizeDanglingTrailingStrongListClosers = true;
        options.NormalizeMetricValueStrongRuns = true;
    }

    private static void ApplyIntelligenceXTranscriptStrict(MarkdownInputNormalizationOptions options) {
        ApplyIntelligenceXTranscript(options);
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
    private const int StrongFlattenMaxIterations = 32;
    private const int LabeledOuterStrongPrefixMaxChars = 120;

    private static readonly Regex ZeroWidthWhitespaceRegex = new Regex(
        @"[\u200B\u2060\uFEFF]",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex EmojiWordJoinRegex = new Regex(
        @"([✅☑✔❌⚠🔥])(?!\s)(?=[\p{L}\p{N}])",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex NumberedChoiceJoinRegex = new Regex(
        @"(\bor|\band|[,;:])(?!\s)(?=\d+\))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex LetterToNumberedChoiceJoinRegex = new Regex(
        @"(?<=[A-Za-z])(?=\d+\))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex SentenceCollapsedBulletRegex = new Regex(
        @"(?<=[\.\!\?\)\]])\s*(?=-\s*(?:\*\*[^\r\n]|[A-Z]{2,}\d+\b))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

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

    private static readonly Regex WrappedSignalFlowLineRegex = new Regex(
        @"(?m)^(?<prefix>\s*-\s+[^\r\n]*?)\*\*(?<inner>[^\r\n]*->\s*\*\*[^\r\n]*?)\*\*(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex SignalFlowPlainLabelTightSpacingRegex = new Regex(
        @"^(?<label>[\p{L}][^:\r\n]{0,120}:)(?<next>[^\s/\\])",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex StatusCollapsedLineRegex = new Regex(
        @"(?m)^(?<lead>\s*\*\*Status:[^\r\n]*?)[ \t]-[ \t](?<rest>\*\*.*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex BulletCollapsedLineRegex = new Regex(
        @"(?m)^(?<lead>\s*-\s[^\r\n]*?)[ \t]-(?<rest>\s*\*{1,2}.*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LegacyStatusSummaryRegex = new Regex(
        @"(?m)^(?<indent>\s*)\*\*Status:\s*(?<value>[^*\r\n]+)\*\*\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LegacyBoldMetricBulletRegex = new Regex(
        @"(?m)^(?<indent>\s*-\s)\*\*(?<label>[^*\r\n:]+):\*\*\s*(?<value>[^\r\n]*?)\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LineStartHostLabelBulletRegex = new Regex(
        @"(?m)^(?<indent>\s*)-(?=[A-Z]{2,}\d+\b)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LineStartMissingSpaceBeforeBoldBulletRegex = new Regex(
        @"(?m)^(?<indent>\s*)-(?=\*\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LineStartUnicodeDashBulletRegex = new Regex(
        @"(?m)^(?<indent>\s*)[‐‑‒–—−](?=(?:\s*\*\*|[A-Z]{2,}\d+\b|[\p{Lu}][\p{L}\p{N}]{1,}\b))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LineStartBoldBulletStrongOpenWhitespaceRegex = new Regex(
        @"(?m)^(?<lead>\s*-\s+\*\*)[ \t]+(?=[^\s*\r\n])",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex SingleStarMetricBulletRegex = new Regex(
        @"(?m)^(?<indent>\s*)-\s*\*(?=[A-Za-z][^\r\n]*:\*\*)",
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

    private static readonly Regex StandaloneSingleHashSeparatorRegex = new Regex(
        @"^\s*#\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex BrokenTwoLineStrongLeadInRegex = new Regex(
        @"^(?<indent>\s*)\*\*(?<label>Result)\s*$",
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

    private static readonly Regex CollapsedOrderedListAfterParenRegex = new Regex(
        @"(?<=\))\s*(?=\d+\.(?:\^\s*|\s*[*_]{2}|\s+)\S)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex CollapsedOrderedListAfterDetailRegex = new Regex(
        @"(?<=\))\s+(?=\d+[.)]\s*[*_]{0,2}\s*\S)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex CollapsedOrderedListAfterStrongRegex = new Regex(
        @"(?<=\*\*)\s+(?=\d+[.)]\s*[*_]{0,2}\s*\S)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex OrderedItemStrongMissingCloseBeforeParenRegex = new Regex(
        @"(?m)^(?<lead>\s*\d+\.\s+)\*\*(?<title>[^*\r\n()]+)\((?<detail>[^)\r\n]+)\)\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex TightParentheticalSpacingRegex = new Regex(
        @"(?:(?<=\*\*)|(?<=[\p{L}\p{N}\)]))\((?=[\p{L}][^\r\n)]*\))",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex NestedStrongSpanRegex = new Regex(
        @"(?<!\S)\*\*(?<left>[^*\r\n]{6,}?\s)\*\*(?<inner>[A-Za-z0-9`][^*:\r\n]*?)\*\*(?<right>[^*\r\n]*?)\*\*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex SimpleNestedStrongSpanRegex = new Regex(
        @"\*\*(?<inner>[^*\r\n]+)\*\*",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex LabeledOuterStrongLineRegex = new Regex(
        @"(?m)^(?<prefix>\s*-\s+[^*\r\n]{2," + LabeledOuterStrongPrefixMaxChars.ToString() + @"}\s+\*\*)(?<body>[^\r\n]*)(?<suffix>\*\*)(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex OrderedListLeadRegex = new Regex(
        @"^\d+[.)]\s+",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex StandaloneHostLabelBulletRegex = new Regex(
        @"^\s*-(?:\s*\*\*)?\s*[A-Z]{2,}\d+(?:\s*\*\*)?\s*:?\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex StructuralMarkdownLineRegex = new Regex(
        @"^(?:[-+*]\s+|\d+[.)]\s+|#{1,6}\s+|>\s?|```|~~~|\|)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex TrailingDanglingStrongListTokenRegex = new Regex(
        @"(?<token>[\p{L}\p{N}_./:-]+)\*{4}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex OveropenedMetricValueStrongRegex = new Regex(
        @"^(?<prefix>\s*(?:-\s+|\d+\.\s+)[^\r\n*]+?\s)\*{4,}(?<value>[^\s*\r\n][^*\r\n]*?)\*{2}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex AdjacentMetricStrongValueRegex = new Regex(
        @"^(?<prefix>\s*(?:-\s+|\d+\.\s+)[^\r\n*]+?\s)\*\*(?<first>[^*\r\n]+)\*\*\*{2}(?<second>[^\s*\r\n][^*\r\n]*?)\*{2}(?<tail>\s*)$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex MissingTrailingStrongMetricCloseRegex = new Regex(
        @"^(?<prefix>\s*(?:-\s+|\d+\.\s+)[^\r\n*]+?\s)\*\*(?<value>[^\r\n*][^\r\n]*?)(?<!\*)\*(?<tail>\s*)$",
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

        if (options.NormalizeZeroWidthSpacingArtifacts) {
            value = ZeroWidthWhitespaceRegex.Replace(value, string.Empty);
        }

        if (options.NormalizeBrokenTwoLineStrongLeadIns) {
            value = ApplyTransformOutsideFencedCodeBlocks(value, RepairBrokenTwoLineStrongLeadIns);
        }

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

        if (options.NormalizeWrappedSignalFlowStrongRuns) {
            value = ApplyRegexOutsideFencedCodeBlocks(
                value,
                WrappedSignalFlowLineRegex,
                static match => {
                    var inner = match.Groups["inner"].Value;
                    var markerIndex = inner.IndexOf("-> **", StringComparison.Ordinal);
                    if (markerIndex < 0) {
                        markerIndex = inner.IndexOf("->**", StringComparison.Ordinal);
                    }

                    if (markerIndex <= 0) {
                        return match.Value;
                    }

                    var headline = inner.Substring(0, markerIndex).TrimEnd();
                    if (headline.Length == 0) {
                        return match.Value;
                    }

                    var flow = inner.Substring(markerIndex).TrimStart();
                    if (flow.StartsWith("->**", StringComparison.Ordinal)) {
                        flow = "-> **" + flow.Substring(4);
                    }

                    if (!flow.StartsWith("-> **", StringComparison.Ordinal)) {
                        return match.Value;
                    }

                    return match.Groups["prefix"].Value + "**" + headline + "** " + flow + match.Groups["tail"].Value;
                });
        }

        if (options.NormalizeSignalFlowLabelSpacing) {
            value = ApplyTransformOutsideFencedCodeBlocks(value, NormalizeSignalFlowLabelSpacing);
        }

        if (options.NormalizeEmojiWordJoins) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, EmojiWordJoinRegex, static match => match.Groups[1].Value + " ");
        }

        if (options.NormalizeCompactNumberedChoiceBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, NumberedChoiceJoinRegex, static match => match.Groups[1].Value + " ");
            value = ApplyRegexOutsideFencedCodeBlocks(value, LetterToNumberedChoiceJoinRegex, static _ => " ");
        }

        if (options.NormalizeSentenceCollapsedBullets) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, SentenceCollapsedBulletRegex, static _ => "\n", preserveInlineCodeSpans: true);
        }

        if (options.NormalizeCollapsedMetricChains) {
            value = ExpandCollapsedMetricLines(value);
            value = NormalizeLegacyMetricBulletLeads(value);
            value = ConvertLegacyMetricMarkdown(value);
        }

        if (options.NormalizeHostLabelBulletArtifacts) {
            value = ApplyTransformOutsideFencedCodeBlocks(value, NormalizeHostLabelBulletArtifacts);
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

        if (options.NormalizeDanglingTrailingStrongListClosers) {
            value = RepairDanglingTrailingStrongListClosers(value);
        }

        if (options.NormalizeMetricValueStrongRuns) {
            value = RepairMalformedMetricValueStrongRuns(value);
        }

        if (options.NormalizeCollapsedOrderedListBoundaries) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, CollapsedOrderedListAfterDetailRegex, static _ => "\n");
            value = ApplyRegexOutsideFencedCodeBlocks(value, CollapsedOrderedListAfterStrongRegex, static _ => "\n");
            value = ApplyRegexOutsideFencedCodeBlocks(value, CollapsedOrderedListAfterParenRegex, static _ => "\n");
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

        if (options.NormalizeStandaloneHashHeadingSeparators) {
            value = RemoveStandaloneHashSeparatorsBeforeHeadings(value);
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

        if (options.NormalizeOrderedListStrongDetailClosures) {
            value = ApplyRegexOutsideFencedCodeBlocks(value, OrderedItemStrongMissingCloseBeforeParenRegex, static match => {
                var lead = match.Groups["lead"].Value;
                var title = match.Groups["title"].Value.Trim();
                var detail = match.Groups["detail"].Value.Trim();
                return lead + "**" + title + "** (" + detail + ")";
            });
        }

        // Keep a final loose-strong pass after ordered-list detail repair because that step can
        // introduce new boundary whitespace inside reconstructed strong delimiters.
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
        return MarkdownFence.ApplyTransformOutsideFencedCodeBlocks(
            value,
            static segment => MarkdownInlineCode.ApplyTransformPreservingInlineCodeSpans(segment, FlattenNestedStrongSpansPreservingInlineCode));
    }

    private static string FlattenNestedStrongSpansPreservingInlineCode(string input) {
        var current = input ?? string.Empty;
        while (true) {
            var flattened = NestedStrongSpanRegex.Replace(
                current,
                static match =>
                    "**"
                    + match.Groups["left"].Value
                    + match.Groups["inner"].Value
                    + match.Groups["right"].Value
                    + "**");
            if (flattened.Equals(current, StringComparison.Ordinal)) {
                break;
            }

            current = flattened;
        }

        return FlattenLabeledOuterStrongSpans(current);
    }

    private static string FlattenLabeledOuterStrongSpans(string input) {
        if (string.IsNullOrEmpty(input) || input.IndexOf("**", StringComparison.Ordinal) < 0) {
            return input ?? string.Empty;
        }

        return LabeledOuterStrongLineRegex.Replace(input, static match => {
            var body = match.Groups["body"].Value;
            if (body.IndexOf("**", StringComparison.Ordinal) < 0) {
                return match.Value;
            }

            var trimmedBody = body.TrimEnd();
            if (trimmedBody.Length == 0) {
                return match.Value;
            }

            var lastBodyChar = trimmedBody[trimmedBody.Length - 1];
            if (lastBodyChar != '.' && lastBodyChar != '!' && lastBodyChar != '?' && lastBodyChar != ')') {
                return match.Value;
            }

            var cleaned = FlattenNestedStrongMarkers(body);
            if (cleaned.Equals(body, StringComparison.Ordinal)) {
                return match.Value;
            }

            return match.Groups["prefix"].Value + cleaned + match.Groups["suffix"].Value + match.Groups["tail"].Value;
        });
    }

    private static string FlattenNestedStrongMarkers(string input) {
        if (string.IsNullOrEmpty(input) || input.IndexOf("**", StringComparison.Ordinal) < 0) {
            return input ?? string.Empty;
        }

        var current = input;
        for (var i = 0; i < StrongFlattenMaxIterations; i++) {
            var next = SimpleNestedStrongSpanRegex.Replace(
                current,
                match => {
                    var inner = match.Groups["inner"].Value;
                    if (inner.Length == 0) {
                        return inner;
                    }

                    var prefix = string.Empty;
                    var suffix = string.Empty;
                    var start = match.Index;
                    var end = match.Index + match.Length;
                    if (start > 0) {
                        var before = current[start - 1];
                        if (!char.IsWhiteSpace(before) && IsWordLikeChar(before) && IsWordLikeChar(inner[0])) {
                            prefix = " ";
                        }
                    }

                    if (end < current.Length) {
                        var after = current[end];
                        if (!char.IsWhiteSpace(after) && IsWordLikeChar(inner[inner.Length - 1]) && IsWordLikeChar(after)) {
                            suffix = " ";
                        }
                    }

                    return prefix + inner + suffix;
                });

            if (next.Equals(current, StringComparison.Ordinal)) {
                return next;
            }

            current = next;
        }

        return current;
    }

    private static bool IsWordLikeChar(char value) {
        return char.IsLetterOrDigit(value);
    }

    private static string NormalizeSignalFlowLabelSpacing(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf("->", StringComparison.Ordinal) < 0) {
            return text ?? string.Empty;
        }

        var hasCrLf = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var changed = false;

        for (var i = 0; i < lines.Length; i++) {
            var line = lines[i] ?? string.Empty;
            if (line.IndexOf("->", StringComparison.Ordinal) < 0
                || line.IndexOf('`') >= 0) {
                continue;
            }

            var rewritten = NormalizeSignalFlowArrowSegments(line);
            if (rewritten.Equals(line, StringComparison.Ordinal)) {
                continue;
            }

            lines[i] = rewritten;
            changed = true;
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", lines);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static string NormalizeSignalFlowArrowSegments(string line) {
        var segments = line.Split(new[] { "->" }, StringSplitOptions.None);
        if (segments.Length < 2) {
            return line;
        }

        var builder = new StringBuilder(line.Length + 8);
        builder.Append(segments[0]);
        for (var i = 1; i < segments.Length; i++) {
            builder.Append("->");
            builder.Append(NormalizeSignalFlowSegmentLabelSpacing(segments[i]));
        }

        return builder.ToString();
    }

    private static string NormalizeSignalFlowSegmentLabelSpacing(string segment) {
        if (string.IsNullOrEmpty(segment)) {
            return segment ?? string.Empty;
        }

        var start = 0;
        while (start < segment.Length && char.IsWhiteSpace(segment[start])) {
            start++;
        }

        if (start >= segment.Length) {
            return segment;
        }

        var strongNormalized = TryNormalizeLeadingStrongSignalLabel(segment, start);
        if (!strongNormalized.Equals(segment, StringComparison.Ordinal)) {
            return strongNormalized;
        }

        return TryNormalizeLeadingPlainSignalLabel(segment, start);
    }

    private static string TryNormalizeLeadingStrongSignalLabel(string segment, int start) {
        if (start + 1 >= segment.Length || segment[start] != '*' || segment[start + 1] != '*') {
            return segment;
        }

        var close = segment.IndexOf("**", start + 2, segment.Length - (start + 2), StringComparison.Ordinal);
        if (close < 0 || close + 2 >= segment.Length) {
            return segment;
        }

        if (segment[close - 1] != ':') {
            return segment;
        }

        var next = segment[close + 2];
        if (char.IsWhiteSpace(next)) {
            return segment;
        }

        return segment.Insert(close + 2, " ");
    }

    private static string TryNormalizeLeadingPlainSignalLabel(string segment, int start) {
        var candidate = segment.Substring(start);
        var match = SignalFlowPlainLabelTightSpacingRegex.Match(candidate);
        if (!match.Success) {
            return segment;
        }

        return segment.Insert(start + match.Groups["label"].Length, " ");
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

    private static string ExpandCollapsedMetricLines(string text) {
        if (string.IsNullOrEmpty(text)) {
            return text ?? string.Empty;
        }

        var newline = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0 ? "\r\n" : "\n";
        var current = text;

        while (true) {
            var afterStatus = StatusCollapsedLineRegex.Replace(
                current,
                match => match.Groups["lead"].Value + newline + "- " + match.Groups["rest"].Value);

            var afterBullets = BulletCollapsedLineRegex.Replace(
                afterStatus,
                match => match.Groups["lead"].Value + newline + "- " + match.Groups["rest"].Value.TrimStart());

            if (afterBullets == current) {
                return afterBullets;
            }

            current = afterBullets;
        }
    }

    private static string NormalizeLegacyMetricBulletLeads(string text) {
        if (string.IsNullOrEmpty(text)) {
            return text ?? string.Empty;
        }

        var spaced = LineStartMissingSpaceBeforeBoldBulletRegex.Replace(text, "${indent}- ");
        return SingleStarMetricBulletRegex.Replace(spaced, "${indent}- **");
    }

    private static string ConvertLegacyMetricMarkdown(string text) {
        if (string.IsNullOrEmpty(text)) {
            return text ?? string.Empty;
        }

        var statusNormalized = LegacyStatusSummaryRegex.Replace(
            text,
            match => {
                var indent = match.Groups["indent"].Value;
                var value = match.Groups["value"].Value.Trim();
                return value.Length == 0 ? indent + "Status" : indent + "Status **" + value + "**";
            });

        return LegacyBoldMetricBulletRegex.Replace(
            statusNormalized,
            match => {
                var indent = match.Groups["indent"].Value;
                var label = match.Groups["label"].Value.Trim();
                var value = match.Groups["value"].Value.Trim();
                if (value.Length == 0) {
                    return indent + label;
                }

                if (value.IndexOf("**", StringComparison.Ordinal) >= 0
                    || value.IndexOf('`') >= 0
                    || value.IndexOf("~~", StringComparison.Ordinal) >= 0
                    || value.IndexOf("==", StringComparison.Ordinal) >= 0) {
                    return indent + label + " " + value;
                }

                return indent + label + " **" + value + "**";
            });
    }

    private static string RepairBrokenTwoLineStrongLeadIns(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf("**", StringComparison.Ordinal) < 0) {
            return text ?? string.Empty;
        }

        var hasCrLf = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var rewritten = new List<string>(lines.Length);
        var changed = false;

        for (var i = 0; i < lines.Length; i++) {
            var current = lines[i] ?? string.Empty;
            if (i + 1 < lines.Length) {
                var currentMatch = BrokenTwoLineStrongLeadInRegex.Match(current);
                if (currentMatch.Success) {
                    var next = lines[i + 1] ?? string.Empty;
                    var closingIndex = next.IndexOf("**", StringComparison.Ordinal);
                    if (closingIndex > 0) {
                        var label = currentMatch.Groups["label"].Value.Trim().TrimEnd(':');
                        var body = next.Substring(0, closingIndex).Trim();
                        var tail = next.Substring(closingIndex + 2).Trim();
                        if (label.Length > 0
                            && body.Length > 0
                            && !StructuralMarkdownLineRegex.IsMatch(body)) {
                            var merged = currentMatch.Groups["indent"].Value
                                         + "**" + label + ":** "
                                         + body
                                         + (tail.Length == 0 ? string.Empty : " " + tail);
                            rewritten.Add(merged);
                            changed = true;
                            i++;
                            continue;
                        }
                    }
                }
            }

            rewritten.Add(current);
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", rewritten);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static string NormalizeHostLabelBulletArtifacts(string text) {
        if (string.IsNullOrEmpty(text)) {
            return text ?? string.Empty;
        }

        var spaced = LineStartUnicodeDashBulletRegex.Replace(text, "${indent}-");
        spaced = LineStartMissingSpaceBeforeBoldBulletRegex.Replace(spaced, "${indent}- ");
        spaced = LineStartBoldBulletStrongOpenWhitespaceRegex.Replace(spaced, "${lead}");
        spaced = LineStartHostLabelBulletRegex.Replace(spaced, "${indent}- ");
        if (spaced.IndexOf('\n') < 0 && spaced.IndexOf('\r') < 0) {
            return spaced;
        }

        var hasCrLf = spaced.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = spaced.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        if (lines.Length < 2) {
            return spaced;
        }

        var merged = new List<string>(lines.Length);
        var changed = !spaced.Equals(text, StringComparison.Ordinal);
        for (var i = 0; i < lines.Length; i++) {
            var current = lines[i] ?? string.Empty;
            if (i + 1 < lines.Length
                && StandaloneHostLabelBulletRegex.IsMatch(current)
                && ShouldAttachHostLabelContinuation(lines[i + 1])) {
                var next = (lines[i + 1] ?? string.Empty).TrimStart();
                merged.Add(current.TrimEnd() + " " + next);
                changed = true;
                i++;
                continue;
            }

            merged.Add(current);
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", merged);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static string RemoveStandaloneHashSeparatorsBeforeHeadings(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf('#') < 0) {
            return text ?? string.Empty;
        }

        var hasCrLf = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var rewritten = new List<string>(lines.Length);
        var changed = false;

        for (var i = 0; i < lines.Length; i++) {
            var current = lines[i] ?? string.Empty;
            if (StandaloneSingleHashSeparatorRegex.IsMatch(current)
                && TryFindNextNonEmptyLine(lines, i + 1, out var nextIndex)
                && IsMarkdownHeadingLine(lines[nextIndex] ?? string.Empty)) {
                changed = true;
                i = nextIndex - 1;
                continue;
            }

            rewritten.Add(current);
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", rewritten);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static string RepairDanglingTrailingStrongListClosers(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf("****", StringComparison.Ordinal) < 0) {
            return text ?? string.Empty;
        }

        var hasCrLf = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var changed = false;

        for (var i = 0; i < lines.Length; i++) {
            var line = lines[i] ?? string.Empty;
            var trimmedStart = line.TrimStart();
            if (!trimmedStart.StartsWith("- ", StringComparison.Ordinal)
                && !OrderedListLeadRegex.IsMatch(trimmedStart)) {
                continue;
            }

            var repaired = TrailingDanglingStrongListTokenRegex.Replace(line, static match => {
                var token = match.Groups["token"].Value.Trim();
                if (token.Length == 0 || token.IndexOf("**", StringComparison.Ordinal) >= 0) {
                    return match.Value;
                }

                return "**" + token + "**" + match.Groups["tail"].Value;
            });

            if (repaired.Equals(line, StringComparison.Ordinal)) {
                continue;
            }

            lines[i] = repaired;
            changed = true;
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", lines);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static string RepairMalformedMetricValueStrongRuns(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf("**", StringComparison.Ordinal) < 0) {
            return text ?? string.Empty;
        }

        var hasCrLf = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var changed = false;

        for (var i = 0; i < lines.Length; i++) {
            var line = lines[i] ?? string.Empty;
            if (!TryRepairMalformedMetricValueStrongRunLine(line, out var repaired)
                || repaired.Equals(line, StringComparison.Ordinal)) {
                continue;
            }

            lines[i] = repaired;
            changed = true;
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", lines);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static bool TryRepairMalformedMetricValueStrongRunLine(string line, out string repaired) {
        repaired = line;
        var trimmedStart = line.TrimStart();
        if (!trimmedStart.StartsWith("- ", StringComparison.Ordinal)
            && !OrderedListLeadRegex.IsMatch(trimmedStart)) {
            return false;
        }

        repaired = OveropenedMetricValueStrongRegex.Replace(line, static match => {
            var value = match.Groups["value"].Value.Trim();
            return value.Length == 0
                ? match.Value
                : match.Groups["prefix"].Value + "**" + value + "**" + match.Groups["tail"].Value;
        });

        repaired = AdjacentMetricStrongValueRegex.Replace(repaired, static match => {
            var first = match.Groups["first"].Value.Trim();
            var second = match.Groups["second"].Value.Trim();
            if (first.Length == 0 || second.Length == 0) {
                return match.Value;
            }

            if (IsSymbolOnlyMetricValue(first)) {
                return match.Groups["prefix"].Value + first + " **" + second + "**" + match.Groups["tail"].Value;
            }

            return match.Groups["prefix"].Value
                   + "**"
                   + first
                   + "** **"
                   + second
                   + "**"
                   + match.Groups["tail"].Value;
        });

        repaired = MissingTrailingStrongMetricCloseRegex.Replace(repaired, static match => {
            var value = match.Groups["value"].Value.Trim();
            return value.Length == 0
                ? match.Value
                : match.Groups["prefix"].Value + "**" + value + "**" + match.Groups["tail"].Value;
        });

        return true;
    }

    private static bool IsSymbolOnlyMetricValue(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        foreach (var ch in value) {
            if (char.IsWhiteSpace(ch)) {
                continue;
            }

            if (char.IsLetterOrDigit(ch)) {
                return false;
            }
        }

        return true;
    }

    private static bool IsMarkdownHeadingLine(string line) {
        var trimmed = line.TrimStart();
        if (trimmed.Length < 4 || trimmed[0] != '#') {
            return false;
        }

        var depth = 0;
        while (depth < trimmed.Length && trimmed[depth] == '#') {
            depth++;
        }

        return depth is >= 2 and <= 6
               && depth < trimmed.Length
               && char.IsWhiteSpace(trimmed[depth]);
    }

    private static bool TryFindNextNonEmptyLine(string[] lines, int startIndex, out int index) {
        for (var i = startIndex; i < lines.Length; i++) {
            if (!string.IsNullOrWhiteSpace(lines[i])) {
                index = i;
                return true;
            }
        }

        index = -1;
        return false;
    }

    private static bool ShouldAttachHostLabelContinuation(string line) {
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        var trimmed = line.TrimStart();
        return !StructuralMarkdownLineRegex.IsMatch(trimmed);
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

            if (MarkdownFence.TryReadContainerAwareFenceRun(line, out _, out var runMarker, out var runLength, out var runSuffix)) {
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

        if (!MarkdownFence.TryReadContainerAwareFenceRun(line, out var linePrefix, out fenceMarker, out fenceRunLength, out var runSuffix)) {
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

        normalizedLine = linePrefix + new string(fenceMarker, fenceRunLength) + language + "\n" + linePrefix + body;
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
