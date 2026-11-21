using System.Text;
// Intentionally avoid heavy regex use; simple scanning is used for resilience and speed.

namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    // Nested containers parse with the same options without front matter, and every nested parse
    // needs the block parser pipeline for its options. Both are made once per parse; the shared
    // options copy is read-only.
    private static MarkdownLoadOptions GetSharedOptionsWithoutFrontMatter(MarkdownLoadOptions source, MarkdownReaderState state) {
        var cache = state.ParseCache;
        if (cache.IsDerived(source, MarkdownReaderDerivedOptions.WithoutFrontMatter)) {
            return source;
        }

        return cache.GetDerived(source, MarkdownReaderDerivedOptions.WithoutFrontMatter)
            ?? cache.AddDerived(source, MarkdownReaderDerivedOptions.WithoutFrontMatter, CloneOptionsWithoutFrontMatter(source));
    }

    // Containers that parse their content with one option changed (footnote bodies, structured
    // table cells, quotes inside list items) reuse one read-only copy per parse, so repeated
    // containers share it and its block parser pipeline.
    private static MarkdownLoadOptions GetSharedDerivedOptions(
        MarkdownLoadOptions source,
        MarkdownReaderState state,
        MarkdownReaderDerivedOptions kind,
        Func<MarkdownLoadOptions, MarkdownLoadOptions> create) {
        var cache = state.ParseCache;
        return cache.GetDerived(source, kind) ?? cache.AddDerived(source, kind, create(source));
    }

    private static MarkdownReaderPipeline GetPipeline(MarkdownLoadOptions options, MarkdownReaderState state) {
        var pipelines = state.ParseCache.Pipelines;
        for (int i = 0; i < pipelines.Count; i++) {
            if (ReferenceEquals(pipelines[i].Key, options)) {
                return pipelines[i].Value;
            }
        }

        var pipeline = MarkdownReaderPipeline.Default(options);
        if (pipelines.Count < MarkdownReaderParseCache.MaxPipelines) {
            pipelines.Add(new KeyValuePair<MarkdownLoadOptions, MarkdownReaderPipeline>(options, pipeline));
        }

        return pipeline;
    }

    private static MarkdownLoadOptions CloneOptionsWithoutFrontMatter(MarkdownLoadOptions source) {
        var clone = new MarkdownLoadOptions {
            FrontMatter = false,
            Callouts = source.Callouts,
            CalloutTitleMode = source.CalloutTitleMode,
            Headings = source.Headings,
            FencedCode = source.FencedCode,
            IndentedCodeBlocks = source.IndentedCodeBlocks,
            Images = source.Images,
            StandaloneImageBlocks = source.StandaloneImageBlocks,
            UnorderedLists = source.UnorderedLists,
            TaskLists = source.TaskLists,
            OrderedLists = source.OrderedLists,
            ListExtras = source.ListExtras,
            StrictListIndentation = source.StrictListIndentation,
            Tables = source.Tables,
            AllowHeaderlessTables = source.AllowHeaderlessTables,
            RequireTableBodyRowPipes = source.RequireTableBodyRowPipes,
            ParseTableCellBlocks = source.ParseTableCellBlocks,
            DefinitionLists = source.DefinitionLists,
            TocPlaceholders = source.TocPlaceholders,
            Footnotes = source.Footnotes,
            SingleTildeStrikethrough = source.SingleTildeStrikethrough,
            Highlight = source.Highlight,
            Inserted = source.Inserted,
            Superscript = source.Superscript,
            Subscript = source.Subscript,
            CjkFriendlyEmphasis = source.CjkFriendlyEmphasis,
            PreferNarrativeSingleLineDefinitions = source.PreferNarrativeSingleLineDefinitions,
            HtmlBlocks = source.HtmlBlocks,
            AllowLooseHtmlBlockStartTags = source.AllowLooseHtmlBlockStartTags,
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
            AutolinkRejectUnderscoreInWwwSubdomainLabels = source.AutolinkRejectUnderscoreInWwwSubdomainLabels,
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
            Math = source.Math,
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
            MaxNestingDepth = source.MaxNestingDepth,
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
        CopyInlineTransformExtensions(source, clone);
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
        var clone = MarkdownReaderState.CreateSharingDefinitionsWith(state);
        clone.ShareParseCacheWith(state);
        clone.SourceLineOffset = state.SourceLineOffset;
        clone.SourceTextMap = state.SourceTextMap;
        clone.SourceLineAbsoluteNumbers = state.SourceLineAbsoluteNumbers;
        clone.ListMarkerIndentOffset = state.ListMarkerIndentOffset;
        clone.CaptureSyntaxTree = state.CaptureSyntaxTree;
        clone.BuildTableCellModels = state.BuildTableCellModels;
        clone.IsListBoundaryProbe = state.IsListBoundaryProbe;
        clone.UseRelativeListMarkerIndent = state.UseRelativeListMarkerIndent;
        clone.SuppressBlockGenericAttributes = state.SuppressBlockGenericAttributes;
        clone.SuppressHeadingGenericAttributes = state.SuppressHeadingGenericAttributes;
        clone.IsMarkdigDefinitionListBody = state.IsMarkdigDefinitionListBody;
        clone.NestedBlockDepth = state.NestedBlockDepth;
        clone.DiscardsNestedListSourceRemap = state.DiscardsNestedListSourceRemap;
        state.LazyQuoteContinuationLines.CopyTo(clone.LazyQuoteContinuationLines);
        state.QuoteContainerLines.CopyTo(clone.QuoteContainerLines);
        state.SuppressedSetextHeadingUnderlineLines.CopyTo(clone.SuppressedSetextHeadingUnderlineLines);
        state.SuppressedParagraphGenericAttributeStartLines.CopyTo(clone.SuppressedParagraphGenericAttributeStartLines);
        return clone;
    }

    private static void CopyFencedBlockExtensions(MarkdownLoadOptions source, MarkdownLoadOptions target) {
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

    private static void CopyBlockParserExtensions(MarkdownLoadOptions source, MarkdownLoadOptions target) {
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

    private static void CopyInlineParserExtensions(MarkdownLoadOptions source, MarkdownLoadOptions target) {
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

    private static void CopyInlineTransformExtensions(MarkdownLoadOptions source, MarkdownLoadOptions target) {
        if (source == null || target == null) {
            return;
        }

        var extensions = source.InlineTransformExtensions;
        target.InlineTransformExtensions.Clear();
        if (extensions.Count == 0) {
            return;
        }

        for (int i = 0; i < extensions.Count; i++) {
            var extension = extensions[i];
            if (extension != null) {
                target.InlineTransformExtensions.Add(extension);
            }
        }
    }

    private static void CopyDocumentTransforms(MarkdownLoadOptions source, MarkdownLoadOptions target) {
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

/// <summary>Values derived from the options of one parse and reused by its nested parses.</summary>
internal sealed class MarkdownReaderParseCache {
    internal const int MaxPipelines = 8;
    private const int MaxDerivedOptions = 8;
    private readonly List<DerivedOptions> _derivedOptions = new List<DerivedOptions>(2);
    internal readonly List<KeyValuePair<MarkdownLoadOptions, MarkdownReaderPipeline>> Pipelines = new List<KeyValuePair<MarkdownLoadOptions, MarkdownReaderPipeline>>(2);

    internal MarkdownLoadOptions? GetDerived(MarkdownLoadOptions source, MarkdownReaderDerivedOptions kind) {
        for (int i = 0; i < _derivedOptions.Count; i++) {
            var entry = _derivedOptions[i];
            if (entry.Kind == kind && ReferenceEquals(entry.Source, source)) {
                return entry.Options;
            }
        }

        return null;
    }

    internal bool IsDerived(MarkdownLoadOptions options, MarkdownReaderDerivedOptions kind) {
        for (int i = 0; i < _derivedOptions.Count; i++) {
            var entry = _derivedOptions[i];
            if (entry.Kind == kind && ReferenceEquals(entry.Options, options)) {
                return true;
            }
        }

        return false;
    }

    internal MarkdownLoadOptions AddDerived(MarkdownLoadOptions source, MarkdownReaderDerivedOptions kind, MarkdownLoadOptions options) {
        if (_derivedOptions.Count < MaxDerivedOptions) {
            _derivedOptions.Add(new DerivedOptions(source, kind, options));
        }

        return options;
    }

    private readonly struct DerivedOptions {
        internal DerivedOptions(MarkdownLoadOptions source, MarkdownReaderDerivedOptions kind, MarkdownLoadOptions options) {
            Source = source;
            Kind = kind;
            Options = options;
        }

        internal MarkdownLoadOptions Source { get; }
        internal MarkdownReaderDerivedOptions Kind { get; }
        internal MarkdownLoadOptions Options { get; }
    }
}

/// <summary>The read-only option copies a parse derives for nested containers.</summary>
internal enum MarkdownReaderDerivedOptions {
    WithoutFrontMatter,
    FootnoteBody,
    TableCellBlocks,
    ListItemQuote
}
