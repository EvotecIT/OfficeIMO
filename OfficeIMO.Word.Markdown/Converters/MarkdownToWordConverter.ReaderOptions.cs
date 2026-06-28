using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Html;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private static Omd.MarkdownReaderOptions CreateEffectiveReaderOptions(MarkdownToWordOptions options) {
            var source = options.ReaderOptions;
            if (source == null) {
                var defaults = new Omd.MarkdownReaderOptions {
                    BaseUri = options.BaseUri,
                    PreferNarrativeSingleLineDefinitions = options.PreferNarrativeSingleLineDefinitions
                };
                WordMarkdownSemanticBlocks.ConfigureReaderOptions(defaults);
                return defaults;
            }

            var effective = new Omd.MarkdownReaderOptions {
                FrontMatter = source.FrontMatter,
                Callouts = source.Callouts,
                Headings = source.Headings,
                FencedCode = source.FencedCode,
                IndentedCodeBlocks = source.IndentedCodeBlocks,
                Images = source.Images,
                UnorderedLists = source.UnorderedLists,
                TaskLists = source.TaskLists,
                OrderedLists = source.OrderedLists,
                Tables = source.Tables,
                DefinitionLists = source.DefinitionLists,
                TocPlaceholders = source.TocPlaceholders,
                Footnotes = source.Footnotes,
                PreferNarrativeSingleLineDefinitions = source.PreferNarrativeSingleLineDefinitions,
                HtmlBlocks = source.HtmlBlocks,
                Paragraphs = source.Paragraphs,
                AutolinkUrls = source.AutolinkUrls,
                AutolinkWwwUrls = source.AutolinkWwwUrls,
                AutolinkWwwScheme = source.AutolinkWwwScheme,
                AutolinkEmails = source.AutolinkEmails,
                BackslashHardBreaks = source.BackslashHardBreaks,
                SoftLineBreaksAsHardLineBreaks = source.SoftLineBreaksAsHardLineBreaks,
                InlineHtml = source.InlineHtml,
                BaseUri = source.BaseUri,
                DisallowScriptUrls = source.DisallowScriptUrls,
                DisallowFileUrls = source.DisallowFileUrls,
                AllowMailtoUrls = source.AllowMailtoUrls,
                AllowDataUrls = source.AllowDataUrls,
                AllowProtocolRelativeUrls = source.AllowProtocolRelativeUrls,
                RestrictUrlSchemes = source.RestrictUrlSchemes,
                AllowedUrlSchemes = source.AllowedUrlSchemes?.ToArray() ?? Array.Empty<string>(),
                InputNormalization = new Omd.MarkdownInputNormalizationOptions {
                    NormalizeSoftWrappedStrongSpans = source.InputNormalization?.NormalizeSoftWrappedStrongSpans ?? false,
                    NormalizeInlineCodeSpanLineBreaks = source.InputNormalization?.NormalizeInlineCodeSpanLineBreaks ?? false,
                    NormalizeEscapedInlineCodeSpans = source.InputNormalization?.NormalizeEscapedInlineCodeSpans ?? false,
                    NormalizeTightStrongBoundaries = source.InputNormalization?.NormalizeTightStrongBoundaries ?? false,
                    NormalizeTightArrowStrongBoundaries = source.InputNormalization?.NormalizeTightArrowStrongBoundaries ?? false,
                    NormalizeBrokenStrongArrowLabels = source.InputNormalization?.NormalizeBrokenStrongArrowLabels ?? false,
                    NormalizeWrappedSignalFlowStrongRuns = source.InputNormalization?.NormalizeWrappedSignalFlowStrongRuns ?? false,
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
                    NormalizeTightParentheticalSpacing = source.InputNormalization?.NormalizeTightParentheticalSpacing ?? false,
                    NormalizeNestedStrongDelimiters = source.InputNormalization?.NormalizeNestedStrongDelimiters ?? false,
                    NormalizeDanglingTrailingStrongListClosers = source.InputNormalization?.NormalizeDanglingTrailingStrongListClosers ?? false,
                    NormalizeMetricValueStrongRuns = source.InputNormalization?.NormalizeMetricValueStrongRuns ?? false
                }
            };

            if (string.IsNullOrWhiteSpace(effective.BaseUri) && !string.IsNullOrWhiteSpace(options.BaseUri)) {
                effective.BaseUri = options.BaseUri;
            }

            if (!effective.PreferNarrativeSingleLineDefinitions && options.PreferNarrativeSingleLineDefinitions) {
                effective.PreferNarrativeSingleLineDefinitions = true;
            }

            CopyBlockParserExtensions(source, effective);
            CopyFencedBlockExtensions(source, effective);
            CopyDocumentTransforms(source, effective);
            WordMarkdownSemanticBlocks.ConfigureReaderOptions(effective);
            return effective;
        }

        private static void CopyFencedBlockExtensions(Omd.MarkdownReaderOptions source, Omd.MarkdownReaderOptions target) {
            if (source.FencedBlockExtensions.Count == 0) {
                return;
            }

            for (int i = 0; i < source.FencedBlockExtensions.Count; i++) {
                var extension = source.FencedBlockExtensions[i];
                if (extension != null) {
                    target.FencedBlockExtensions.Add(extension);
                }
            }
        }

        private static void CopyBlockParserExtensions(Omd.MarkdownReaderOptions source, Omd.MarkdownReaderOptions target) {
            target.BlockParserExtensions.Clear();
            if (source.BlockParserExtensions.Count == 0) {
                return;
            }

            for (int i = 0; i < source.BlockParserExtensions.Count; i++) {
                var extension = source.BlockParserExtensions[i];
                if (extension != null) {
                    target.BlockParserExtensions.Add(extension);
                }
            }
        }

        private static void CopyDocumentTransforms(Omd.MarkdownReaderOptions source, Omd.MarkdownReaderOptions target) {
            if (source.DocumentTransforms.Count == 0) {
                return;
            }

            for (var i = 0; i < source.DocumentTransforms.Count; i++) {
                var transform = source.DocumentTransforms[i];
                target.DocumentTransforms.Add(transform);
            }
        }

    }
}
