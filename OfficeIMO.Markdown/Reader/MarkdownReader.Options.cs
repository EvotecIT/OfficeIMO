using System.IO;
using System.Linq;
using System.Text;
// Intentionally avoid heavy regex use; simple scanning is used for resilience and speed.

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static MarkdownReaderOptions CloneOptionsWithoutFrontMatter(MarkdownReaderOptions source) {
        var clone = new MarkdownReaderOptions {
            FrontMatter = false,
            Callouts = source.Callouts,
            CalloutTitleMode = source.CalloutTitleMode,
            Headings = source.Headings,
            FencedCode = source.FencedCode,
            IndentedCodeBlocks = source.IndentedCodeBlocks,
            Images = source.Images,
            UnorderedLists = source.UnorderedLists,
            TaskLists = source.TaskLists,
            OrderedLists = source.OrderedLists,
            ListExtras = source.ListExtras,
            Tables = source.Tables,
            AllowHeaderlessTables = source.AllowHeaderlessTables,
            ParseTableCellBlocks = source.ParseTableCellBlocks,
            DefinitionLists = source.DefinitionLists,
            TocPlaceholders = source.TocPlaceholders,
            Footnotes = source.Footnotes,
            SingleTildeStrikethrough = source.SingleTildeStrikethrough,
            Subscript = source.Subscript,
            CjkFriendlyEmphasis = source.CjkFriendlyEmphasis,
            PreferNarrativeSingleLineDefinitions = source.PreferNarrativeSingleLineDefinitions,
            HtmlBlocks = source.HtmlBlocks,
            PreserveHtmlBlockBlankLineContent = source.PreserveHtmlBlockBlankLineContent,
            Paragraphs = source.Paragraphs,
            AutolinkUrls = source.AutolinkUrls,
            AutolinkAllowDomainWithoutPeriod = source.AutolinkAllowDomainWithoutPeriod,
            AutolinkAllowQueryAndFragmentSpecialCharacters = source.AutolinkAllowQueryAndFragmentSpecialCharacters,
            AutolinkAllowBalancedParenthesesWithTrailingPunctuation = source.AutolinkAllowBalancedParenthesesWithTrailingPunctuation,
            AutolinkAllowTrailingPunctuationBeforeClosingParenthesis = source.AutolinkAllowTrailingPunctuationBeforeClosingParenthesis,
            AutolinkTrimSingleTrailingPunctuationOrUnderscore = source.AutolinkTrimSingleTrailingPunctuationOrUnderscore,
            AutolinkKeepTrailingSemicolonPunctuation = source.AutolinkKeepTrailingSemicolonPunctuation,
            AutolinkRequireLowercaseWwwPrefix = source.AutolinkRequireLowercaseWwwPrefix,
            AutolinkRejectUnderscoreInWwwHost = source.AutolinkRejectUnderscoreInWwwHost,
            AutolinkRejectUnderscoreInUrlHost = source.AutolinkRejectUnderscoreInUrlHost,
            AutolinkRejectUserInfoAuthority = source.AutolinkRejectUserInfoAuthority,
            AutolinkAllowClosingBracketInUrl = source.AutolinkAllowClosingBracketInUrl,
            AutolinkKeepTrailingQuotePunctuation = source.AutolinkKeepTrailingQuotePunctuation,
            AutolinkRequireLowercaseBareSchemePrefix = source.AutolinkRequireLowercaseBareSchemePrefix,
            AutolinkBareMailtoDisplayAddressOnly = source.AutolinkBareMailtoDisplayAddressOnly,
            AutolinkBareMailtoMarkdigSemicolonHandling = source.AutolinkBareMailtoMarkdigSemicolonHandling,
            AutolinkValidPreviousCharacters = source.AutolinkValidPreviousCharacters,
            AutolinkBareSchemeUrls = source.AutolinkBareSchemeUrls,
            AutolinkBareSchemePrefixes = source.AutolinkBareSchemePrefixes == null
                ? null
                : (string[])source.AutolinkBareSchemePrefixes.Clone(),
            AutolinkWwwUrls = source.AutolinkWwwUrls,
            AutolinkWwwScheme = source.AutolinkWwwScheme,
            AutolinkEmails = source.AutolinkEmails,
            BackslashHardBreaks = source.BackslashHardBreaks,
            SoftLineBreaksAsHardLineBreaks = source.SoftLineBreaksAsHardLineBreaks,
            InlineHtml = source.InlineHtml,
            Abbreviations = source.Abbreviations,
            GenericAttributes = source.GenericAttributes,
            CustomContainers = source.CustomContainers,
            BaseUri = source.BaseUri,
            DisallowScriptUrls = source.DisallowScriptUrls,
            DisallowFileUrls = source.DisallowFileUrls,
            AllowMailtoUrls = source.AllowMailtoUrls,
            AllowDataUrls = source.AllowDataUrls,
            AllowProtocolRelativeUrls = source.AllowProtocolRelativeUrls,
            RestrictUrlSchemes = source.RestrictUrlSchemes,
            AllowedUrlSchemes = source.AllowedUrlSchemes,
            PreserveTrivia = source.PreserveTrivia,
            MaxInputCharacters = source.MaxInputCharacters,
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeZeroWidthSpacingArtifacts = source.InputNormalization?.NormalizeZeroWidthSpacingArtifacts ?? false,
                NormalizeEmojiWordJoins = source.InputNormalization?.NormalizeEmojiWordJoins ?? false,
                NormalizeCompactNumberedChoiceBoundaries = source.InputNormalization?.NormalizeCompactNumberedChoiceBoundaries ?? false,
                NormalizeSentenceCollapsedBullets = source.InputNormalization?.NormalizeSentenceCollapsedBullets ?? false,
                NormalizeSoftWrappedStrongSpans = source.InputNormalization?.NormalizeSoftWrappedStrongSpans ?? false,
                NormalizeInlineCodeSpanLineBreaks = source.InputNormalization?.NormalizeInlineCodeSpanLineBreaks ?? false,
                NormalizeEscapedInlineCodeSpans = source.InputNormalization?.NormalizeEscapedInlineCodeSpans ?? false,
                NormalizeTightStrongBoundaries = source.InputNormalization?.NormalizeTightStrongBoundaries ?? false,
                NormalizeTightArrowStrongBoundaries = source.InputNormalization?.NormalizeTightArrowStrongBoundaries ?? false,
                NormalizeBrokenStrongArrowLabels = source.InputNormalization?.NormalizeBrokenStrongArrowLabels ?? false,
                NormalizeWrappedSignalFlowStrongRuns = source.InputNormalization?.NormalizeWrappedSignalFlowStrongRuns ?? false,
                NormalizeSignalFlowLabelSpacing = source.InputNormalization?.NormalizeSignalFlowLabelSpacing ?? false,
                NormalizeCollapsedMetricChains = source.InputNormalization?.NormalizeCollapsedMetricChains ?? false,
                NormalizeHostLabelBulletArtifacts = source.InputNormalization?.NormalizeHostLabelBulletArtifacts ?? false,
                NormalizeTightColonSpacing = source.InputNormalization?.NormalizeTightColonSpacing ?? false,
                NormalizeHeadingListBoundaries = source.InputNormalization?.NormalizeHeadingListBoundaries ?? false,
                NormalizeCompactStrongLabelListBoundaries = source.InputNormalization?.NormalizeCompactStrongLabelListBoundaries ?? false,
                NormalizeCompactHeadingBoundaries = source.InputNormalization?.NormalizeCompactHeadingBoundaries ?? false,
                NormalizeStandaloneHashHeadingSeparators = source.InputNormalization?.NormalizeStandaloneHashHeadingSeparators ?? false,
                NormalizeBrokenTwoLineStrongLeadIns = source.InputNormalization?.NormalizeBrokenTwoLineStrongLeadIns ?? false,
                NormalizeColonListBoundaries = source.InputNormalization?.NormalizeColonListBoundaries ?? false,
                NormalizeCompactFenceBodyBoundaries = source.InputNormalization?.NormalizeCompactFenceBodyBoundaries ?? false,
                NormalizeLooseStrongDelimiters = source.InputNormalization?.NormalizeLooseStrongDelimiters ?? false,
                NormalizeOrderedListMarkerSpacing = source.InputNormalization?.NormalizeOrderedListMarkerSpacing ?? false,
                NormalizeOrderedListParenMarkers = source.InputNormalization?.NormalizeOrderedListParenMarkers ?? false,
                NormalizeOrderedListCaretArtifacts = source.InputNormalization?.NormalizeOrderedListCaretArtifacts ?? false,
                NormalizeCollapsedOrderedListBoundaries = source.InputNormalization?.NormalizeCollapsedOrderedListBoundaries ?? false,
                NormalizeOrderedListStrongDetailClosures = source.InputNormalization?.NormalizeOrderedListStrongDetailClosures ?? false,
                NormalizeTightParentheticalSpacing = source.InputNormalization?.NormalizeTightParentheticalSpacing ?? false,
                NormalizeNestedStrongDelimiters = source.InputNormalization?.NormalizeNestedStrongDelimiters ?? false,
                NormalizeDanglingTrailingStrongListClosers = source.InputNormalization?.NormalizeDanglingTrailingStrongListClosers ?? false,
                NormalizeMetricValueStrongRuns = source.InputNormalization?.NormalizeMetricValueStrongRuns ?? false
            }
        };

        CopyBlockParserExtensions(source, clone);
        CopyInlineParserExtensions(source, clone);
        CopyFencedBlockExtensions(source, clone);
        CopyDocumentTransforms(source, clone);
        return clone;
    }

    private static MarkdownInputNormalizationOptions? CreatePreParseNormalizationOptions(MarkdownInputNormalizationOptions? source) {
        bool normalizeZeroWidthSpacingArtifacts = source?.NormalizeZeroWidthSpacingArtifacts ?? false;
        bool normalizeEmojiWordJoins = source?.NormalizeEmojiWordJoins ?? false;
        bool normalizeCompactNumberedChoiceBoundaries = source?.NormalizeCompactNumberedChoiceBoundaries ?? false;
        bool normalizeSentenceCollapsedBullets = source?.NormalizeSentenceCollapsedBullets ?? false;
        bool normalizeSoftWrappedStrong = source?.NormalizeSoftWrappedStrongSpans ?? false;
        bool normalizeInlineCodeLineBreaks = source?.NormalizeInlineCodeSpanLineBreaks ?? false;
        bool normalizeLooseStrongDelimiters = source?.NormalizeLooseStrongDelimiters ?? false;
        bool normalizeTightArrowStrongBoundaries = source?.NormalizeTightArrowStrongBoundaries ?? false;
        bool normalizeBrokenStrongArrowLabels = source?.NormalizeBrokenStrongArrowLabels ?? false;
        // These repairs stay on the text side because malformed input would otherwise parse
        // into the wrong block/inline structure. Recoverable paragraph/heading/list boundary
        // cleanup is intentionally excluded here and handled later via built-in document
        // transforms so the reader can normalize from the AST whenever markdown is already
        // parseable.
        bool normalizeWrappedSignalFlowStrongRuns = source?.NormalizeWrappedSignalFlowStrongRuns ?? false;
        bool normalizeSignalFlowLabelSpacing = source?.NormalizeSignalFlowLabelSpacing ?? false;
        bool normalizeCollapsedMetricChains = source?.NormalizeCollapsedMetricChains ?? false;
        bool normalizeHostLabelBulletArtifacts = source?.NormalizeHostLabelBulletArtifacts ?? false;
        bool normalizeBrokenTwoLineStrongLeadIns = source?.NormalizeBrokenTwoLineStrongLeadIns ?? false;
        bool normalizeCompactFenceBodyBoundaries = source?.NormalizeCompactFenceBodyBoundaries ?? false;
        bool normalizeOrderedListMarkerSpacing = source?.NormalizeOrderedListMarkerSpacing ?? false;
        bool normalizeOrderedListParenMarkers = source?.NormalizeOrderedListParenMarkers ?? false;
        bool normalizeOrderedListCaretArtifacts = source?.NormalizeOrderedListCaretArtifacts ?? false;
        bool normalizeCollapsedOrderedListBoundaries = source?.NormalizeCollapsedOrderedListBoundaries ?? false;
        bool normalizeOrderedListStrongDetailClosures = source?.NormalizeOrderedListStrongDetailClosures ?? false;
        bool normalizeNestedStrongDelimiters = source?.NormalizeNestedStrongDelimiters ?? false;

        if (!normalizeZeroWidthSpacingArtifacts
            && !normalizeEmojiWordJoins
            && !normalizeCompactNumberedChoiceBoundaries
            && !normalizeSentenceCollapsedBullets
            && !normalizeSoftWrappedStrong
            && !normalizeInlineCodeLineBreaks
            && !normalizeLooseStrongDelimiters
            && !normalizeTightArrowStrongBoundaries
            && !normalizeBrokenStrongArrowLabels
            && !normalizeWrappedSignalFlowStrongRuns
            && !normalizeSignalFlowLabelSpacing
            && !normalizeCollapsedMetricChains
            && !normalizeHostLabelBulletArtifacts
            && !normalizeBrokenTwoLineStrongLeadIns
            && !normalizeCompactFenceBodyBoundaries
            && !normalizeOrderedListMarkerSpacing
            && !normalizeOrderedListParenMarkers
            && !normalizeOrderedListCaretArtifacts
            && !normalizeCollapsedOrderedListBoundaries
            && !normalizeOrderedListStrongDetailClosures
            && !normalizeNestedStrongDelimiters) {
            return null;
        }

        return new MarkdownInputNormalizationOptions {
            NormalizeZeroWidthSpacingArtifacts = normalizeZeroWidthSpacingArtifacts,
            NormalizeEmojiWordJoins = normalizeEmojiWordJoins,
            NormalizeCompactNumberedChoiceBoundaries = normalizeCompactNumberedChoiceBoundaries,
            NormalizeSentenceCollapsedBullets = normalizeSentenceCollapsedBullets,
            NormalizeSoftWrappedStrongSpans = normalizeSoftWrappedStrong,
            NormalizeInlineCodeSpanLineBreaks = normalizeInlineCodeLineBreaks,
            NormalizeLooseStrongDelimiters = normalizeLooseStrongDelimiters,
            NormalizeTightArrowStrongBoundaries = normalizeTightArrowStrongBoundaries,
            NormalizeBrokenStrongArrowLabels = normalizeBrokenStrongArrowLabels,
            NormalizeWrappedSignalFlowStrongRuns = normalizeWrappedSignalFlowStrongRuns,
            NormalizeSignalFlowLabelSpacing = normalizeSignalFlowLabelSpacing,
            NormalizeCollapsedMetricChains = normalizeCollapsedMetricChains,
            NormalizeHostLabelBulletArtifacts = normalizeHostLabelBulletArtifacts,
            NormalizeBrokenTwoLineStrongLeadIns = normalizeBrokenTwoLineStrongLeadIns,
            NormalizeCompactFenceBodyBoundaries = normalizeCompactFenceBodyBoundaries,
            NormalizeOrderedListMarkerSpacing = normalizeOrderedListMarkerSpacing,
            NormalizeOrderedListParenMarkers = normalizeOrderedListParenMarkers,
            NormalizeOrderedListCaretArtifacts = normalizeOrderedListCaretArtifacts,
            NormalizeCollapsedOrderedListBoundaries = normalizeCollapsedOrderedListBoundaries,
            NormalizeOrderedListStrongDetailClosures = normalizeOrderedListStrongDetailClosures,
            NormalizeNestedStrongDelimiters = normalizeNestedStrongDelimiters
        };
    }

    private static MarkdownReaderState CloneState(MarkdownReaderState state) {
        var clone = new MarkdownReaderState();
        foreach (var kvp in state.LinkRefs) clone.LinkRefs[kvp.Key] = kvp.Value;
        foreach (var kvp in state.Abbreviations) clone.Abbreviations[kvp.Key] = kvp.Value;
        clone.SourceLineOffset = state.SourceLineOffset;
        clone.SourceTextMap = state.SourceTextMap;
        clone.SourceLineAbsoluteNumbers = state.SourceLineAbsoluteNumbers;
        clone.ListMarkerIndentOffset = state.ListMarkerIndentOffset;
        clone.SuppressBlockGenericAttributes = state.SuppressBlockGenericAttributes;
        clone.SuppressHeadingGenericAttributes = state.SuppressHeadingGenericAttributes;
        clone.IsMarkdigDefinitionListBody = state.IsMarkdigDefinitionListBody;
        foreach (var line in state.LazyQuoteContinuationLines) clone.LazyQuoteContinuationLines.Add(line);
        foreach (var line in state.QuoteContainerLines) clone.QuoteContainerLines.Add(line);
        foreach (var line in state.SuppressedSetextHeadingUnderlineLines) clone.SuppressedSetextHeadingUnderlineLines.Add(line);
        foreach (var line in state.SuppressedParagraphGenericAttributeStartLines) clone.SuppressedParagraphGenericAttributeStartLines.Add(line);
        return clone;
    }

    private static void CopyFencedBlockExtensions(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        var extensions = source.FencedBlockExtensions;
        if (extensions.Count == 0) {
            return;
        }

        for (int i = 0; i < extensions.Count; i++) {
            var extension = extensions[i];
            if (extension != null) {
                target.FencedBlockExtensions.Add(extension);
            }
        }
    }

    private static void CopyBlockParserExtensions(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        if (source == null || target == null) {
            return;
        }

        var extensions = source.BlockParserExtensions;
        target.BlockParserExtensions.Clear();
        if (extensions.Count == 0) {
            return;
        }

        for (int i = 0; i < extensions.Count; i++) {
            var extension = extensions[i];
            if (extension != null) {
                target.BlockParserExtensions.Add(extension);
            }
        }
    }

    private static void CopyInlineParserExtensions(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        if (source == null || target == null) {
            return;
        }

        var extensions = source.InlineParserExtensions;
        target.InlineParserExtensions.Clear();
        if (extensions.Count == 0) {
            return;
        }

        for (int i = 0; i < extensions.Count; i++) {
            var extension = extensions[i];
            if (extension != null) {
                target.InlineParserExtensions.Add(extension);
            }
        }
    }

    private static void CopyDocumentTransforms(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        if (source == null || target == null) {
            return;
        }

        var transforms = source.DocumentTransforms;
        if (transforms.Count == 0) {
            return;
        }

        for (int i = 0; i < transforms.Count; i++) {
            var transform = transforms[i];
            target.DocumentTransforms.Add(transform);
        }
    }
}
