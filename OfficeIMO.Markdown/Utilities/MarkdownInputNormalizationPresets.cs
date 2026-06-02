namespace OfficeIMO.Markdown;

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
