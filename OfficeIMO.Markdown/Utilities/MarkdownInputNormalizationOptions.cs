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
    /// <see cref="MarkdownReader"/> prefers to apply this via a document transform after parse when the markdown
    /// already parsed into a recoverable heading block.
    /// Default: false.
    /// </summary>
    public bool NormalizeHeadingListBoundaries { get; set; } = false;

    /// <summary>
    /// When true, inserts a missing newline before compact unordered strong-label list markers
    /// that were emitted inline after punctuation or symbol characters
    /// (for example, <c>✅- **FSMO:** ok</c> becomes <c>✅\n- **FSMO:** ok</c>).
    /// <see cref="MarkdownReader"/> prefers to apply this via a document transform after parse when the markdown
    /// already parsed into a recoverable paragraph or simple unordered list-item structure.
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
    /// <see cref="MarkdownReader"/> prefers to apply this via a document transform after parse when the markdown
    /// already parsed into a recoverable paragraph block.
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
