using System;
using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.MarkdownRenderer.IntelligenceX;
using Xunit;

namespace OfficeIMO.Tests {
    public class Markdown_IntelligenceXTranscript_Tests {
        [Fact]
        public void HtmlStyle_ChatAuto_EmitsChatMarkerAndAutoThemeCss() {
            var doc = MarkdownReader.Parse("Hello");
            var parts = doc.ToHtmlParts(new HtmlOptions { Kind = HtmlKind.Fragment, Style = HtmlStyle.ChatAuto });

            Assert.Contains("omd-chat", parts.Css);
            Assert.Contains("prefers-color-scheme", parts.Css);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscript_UsesChatStyleAndScopedCss() {
            var opts = MarkdownRendererPresets.CreateIntelligenceXTranscript();
            Assert.Equal(HtmlStyle.ChatAuto, opts.HtmlOptions.Style);
            Assert.Equal("#omdRoot article.markdown-body", opts.HtmlOptions.CssScopeSelector);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateStrict_UsesGenericStyleAndLeavesChatChromeOff() {
            var opts = MarkdownRendererPresets.CreateStrict();

            Assert.Equal(HtmlStyle.GithubAuto, opts.HtmlOptions.Style);
            Assert.Equal("article.markdown-body", opts.HtmlOptions.CssScopeSelector);
            Assert.False(opts.EnableCodeCopyButtons);
            Assert.False(opts.EnableTableCopyButtons);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateStrictPortable_Disables_OfficeImoOnly_Reader_Extensions() {
            var opts = MarkdownRendererPresets.CreateStrictPortable();

            Assert.Equal(HtmlStyle.GithubAuto, opts.HtmlOptions.Style);
            Assert.Equal("article.markdown-body", opts.HtmlOptions.CssScopeSelector);
            Assert.False(opts.ReaderOptions.Callouts);
            Assert.False(opts.ReaderOptions.TaskLists);
            Assert.False(opts.ReaderOptions.TocPlaceholders);
            Assert.False(opts.ReaderOptions.Footnotes);
            Assert.False(opts.ReaderOptions.AutolinkUrls);
            Assert.False(opts.ReaderOptions.AutolinkWwwUrls);
            Assert.False(opts.ReaderOptions.AutolinkEmails);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscriptPortable_Disables_OfficeImoOnly_Reader_Extensions() {
            var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptPortable();

            Assert.Equal(HtmlStyle.ChatAuto, opts.HtmlOptions.Style);
            Assert.Equal("#omdRoot article.markdown-body", opts.HtmlOptions.CssScopeSelector);
            Assert.False(opts.ReaderOptions.Callouts);
            Assert.False(opts.ReaderOptions.TaskLists);
            Assert.False(opts.ReaderOptions.TocPlaceholders);
            Assert.False(opts.ReaderOptions.Footnotes);
            Assert.False(opts.ReaderOptions.AutolinkUrls);
            Assert.False(opts.ReaderOptions.AutolinkWwwUrls);
            Assert.False(opts.ReaderOptions.AutolinkEmails);
            Assert.NotEmpty(opts.HtmlOptions.BlockRenderExtensions);
            Assert.NotNull(opts.HtmlOptions.TocHtmlRenderer);
            Assert.NotNull(opts.HtmlOptions.FootnoteSectionHtmlRenderer);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateStrict_Can_Target_CommonMark_Profile() {
            var opts = MarkdownRendererPresets.CreateStrict(MarkdownReaderOptions.MarkdownDialectProfile.CommonMark);

            Assert.Equal(HtmlStyle.GithubAuto, opts.HtmlOptions.Style);
            Assert.False(opts.ReaderOptions.FrontMatter);
            Assert.False(opts.ReaderOptions.Callouts);
            Assert.False(opts.ReaderOptions.TaskLists);
            Assert.False(opts.ReaderOptions.Tables);
            Assert.False(opts.ReaderOptions.DefinitionLists);
            Assert.False(opts.ReaderOptions.TocPlaceholders);
            Assert.False(opts.ReaderOptions.Footnotes);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscript_Can_Target_Gfm_Profile() {
            var opts = MarkdownRendererPresets.CreateIntelligenceXTranscript(MarkdownReaderOptions.MarkdownDialectProfile.GitHubFlavoredMarkdown);

            Assert.Equal(HtmlStyle.ChatAuto, opts.HtmlOptions.Style);
            Assert.Equal("#omdRoot article.markdown-body", opts.HtmlOptions.CssScopeSelector);
            Assert.False(opts.ReaderOptions.FrontMatter);
            Assert.False(opts.ReaderOptions.Callouts);
            Assert.True(opts.ReaderOptions.TaskLists);
            Assert.True(opts.ReaderOptions.Tables);
            Assert.False(opts.ReaderOptions.DefinitionLists);
            Assert.False(opts.ReaderOptions.TocPlaceholders);
            Assert.True(opts.ReaderOptions.Footnotes);
            Assert.True(opts.ReaderOptions.AutolinkUrls);
            Assert.True(opts.ReaderOptions.AutolinkWwwUrls);
            Assert.True(opts.ReaderOptions.AutolinkEmails);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscript_BuildsOnStrictDefaults() {
            var strict = MarkdownRendererPresets.CreateStrict();
            var chat = MarkdownRendererPresets.CreateIntelligenceXTranscript();

            Assert.Equal(strict.ReaderOptions.RestrictUrlSchemes, chat.ReaderOptions.RestrictUrlSchemes);
            Assert.Equal(strict.HtmlOptions.BlockExternalHttpImages, chat.HtmlOptions.BlockExternalHttpImages);
            Assert.Equal(strict.MaxMarkdownChars, chat.MaxMarkdownChars);
            Assert.Equal(strict.MaxBodyHtmlBytes, chat.MaxBodyHtmlBytes);
            Assert.True(chat.EnableCodeCopyButtons);
            Assert.True(chat.EnableTableCopyButtons);
        }

        [Fact]
        public void MarkdownRendererPresets_ApplyPortableHtmlOutputProfile_Installs_Portable_Block_Fallbacks() {
            var opts = MarkdownRendererPresets.CreateStrict();

            MarkdownRendererPresets.ApplyPortableHtmlOutputProfile(opts);

            Assert.NotEmpty(opts.HtmlOptions.BlockRenderExtensions);
            Assert.NotNull(opts.HtmlOptions.TocHtmlRenderer);
            Assert.NotNull(opts.HtmlOptions.FootnoteSectionHtmlRenderer);
        }

        [Fact]
        public void MarkdownRendererPresets_ApplyChatPresentation_Can_Compose_Generic_Preset_Into_Chat_Surface() {
            var opts = MarkdownRendererPresets.CreateStrictMinimal();

            MarkdownRendererPresets.ApplyChatPresentation(opts, enableCopyButtons: false);
            MarkdownRendererIntelligenceXAdapter.Apply(opts);
            MarkdownRendererIntelligenceXLegacyMigration.Apply(opts);

            Assert.Equal(HtmlStyle.ChatAuto, opts.HtmlOptions.Style);
            Assert.Equal("#omdRoot article.markdown-body", opts.HtmlOptions.CssScopeSelector);
            Assert.False(opts.EnableCodeCopyButtons);
            Assert.False(opts.EnableTableCopyButtons);
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-network", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-dataview", StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscriptMinimal_Matches_Composed_Generic_Preset() {
            var composed = MarkdownRendererPresets.CreateStrictMinimal();
            var transcriptReader = MarkdownTranscriptPreparation.CreateIntelligenceXTranscriptReaderOptions(
                preservesGroupedDefinitionLikeParagraphs: false,
                visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);

            composed.ReaderOptions = transcriptReader;
            composed.NormalizeZeroWidthSpacingArtifacts = true;
            composed.NormalizeEmojiWordJoins = true;
            composed.NormalizeCompactNumberedChoiceBoundaries = true;
            composed.NormalizeSentenceCollapsedBullets = true;
            composed.NormalizeWrappedSignalFlowStrongRuns = true;
            composed.NormalizeSignalFlowLabelSpacing = true;
            composed.NormalizeCollapsedMetricChains = true;
            composed.NormalizeHostLabelBulletArtifacts = true;
            composed.NormalizeCollapsedOrderedListBoundaries = true;
            composed.NormalizeOrderedListStrongDetailClosures = true;
            composed.NormalizeHeadingListBoundaries = true;
            composed.NormalizeCompactStrongLabelListBoundaries = true;
            composed.NormalizeCompactHeadingBoundaries = true;
            composed.NormalizeColonListBoundaries = true;
            composed.NormalizeStandaloneHashHeadingSeparators = true;
            composed.NormalizeBrokenTwoLineStrongLeadIns = true;
            composed.NormalizeDanglingTrailingStrongListClosers = true;
            composed.NormalizeMetricValueStrongRuns = true;
            MarkdownRendererPresets.ApplyChatPresentation(composed, enableCopyButtons: false);
            MarkdownRendererIntelligenceXAdapter.Apply(composed);
            MarkdownRendererIntelligenceXLegacyMigration.Apply(composed);

            var transcript = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();

        Assert.Equal(transcript.HtmlOptions.Style, composed.HtmlOptions.Style);
        Assert.Equal(transcript.HtmlOptions.CssScopeSelector, composed.HtmlOptions.CssScopeSelector);
        Assert.Equal(transcript.EnableCodeCopyButtons, composed.EnableCodeCopyButtons);
        Assert.Equal(transcript.EnableTableCopyButtons, composed.EnableTableCopyButtons);
        Assert.Equal(transcript.NormalizeZeroWidthSpacingArtifacts, composed.NormalizeZeroWidthSpacingArtifacts);
        Assert.Equal(transcript.NormalizeEmojiWordJoins, composed.NormalizeEmojiWordJoins);
        Assert.Equal(transcript.NormalizeCompactNumberedChoiceBoundaries, composed.NormalizeCompactNumberedChoiceBoundaries);
        Assert.Equal(transcript.NormalizeSentenceCollapsedBullets, composed.NormalizeSentenceCollapsedBullets);
        Assert.Equal(transcript.NormalizeWrappedSignalFlowStrongRuns, composed.NormalizeWrappedSignalFlowStrongRuns);
        Assert.Equal(transcript.NormalizeSignalFlowLabelSpacing, composed.NormalizeSignalFlowLabelSpacing);
        Assert.Equal(transcript.NormalizeCollapsedMetricChains, composed.NormalizeCollapsedMetricChains);
        Assert.Equal(transcript.NormalizeHostLabelBulletArtifacts, composed.NormalizeHostLabelBulletArtifacts);
        Assert.Equal(transcript.NormalizeCollapsedOrderedListBoundaries, composed.NormalizeCollapsedOrderedListBoundaries);
        Assert.Equal(transcript.NormalizeOrderedListStrongDetailClosures, composed.NormalizeOrderedListStrongDetailClosures);
        Assert.Equal(transcript.NormalizeHeadingListBoundaries, composed.NormalizeHeadingListBoundaries);
        Assert.Equal(transcript.NormalizeCompactStrongLabelListBoundaries, composed.NormalizeCompactStrongLabelListBoundaries);
        Assert.Equal(transcript.NormalizeCompactHeadingBoundaries, composed.NormalizeCompactHeadingBoundaries);
        Assert.Equal(transcript.NormalizeColonListBoundaries, composed.NormalizeColonListBoundaries);
            Assert.Equal(transcript.NormalizeStandaloneHashHeadingSeparators, composed.NormalizeStandaloneHashHeadingSeparators);
            Assert.Equal(transcript.NormalizeBrokenTwoLineStrongLeadIns, composed.NormalizeBrokenTwoLineStrongLeadIns);
            Assert.Equal(transcript.NormalizeDanglingTrailingStrongListClosers, composed.NormalizeDanglingTrailingStrongListClosers);
            Assert.Equal(transcript.NormalizeMetricValueStrongRuns, composed.NormalizeMetricValueStrongRuns);
            Assert.Equal(transcript.ReaderOptions.DocumentTransforms.Count, composed.ReaderOptions.DocumentTransforms.Count);
            Assert.Equal(transcript.MarkdownPreProcessors.Count, composed.MarkdownPreProcessors.Count);
            Assert.Equal(
                transcript.FencedCodeBlockRenderers.SelectMany(renderer => renderer.Languages).OrderBy(value => value, StringComparer.OrdinalIgnoreCase),
                composed.FencedCodeBlockRenderers.SelectMany(renderer => renderer.Languages).OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscript_Composes_Definition_Compatibility_Transform() {
            var opts = MarkdownRendererPresets.CreateIntelligenceXTranscript();

            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform => transform is MarkdownSimpleDefinitionListParagraphTransform);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform =>
                transform is MarkdownJsonVisualCodeBlockTransform visual
                && visual.LanguageMode == MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscript_Reuses_Shared_Transcript_Reader_Contract() {
            var expected = MarkdownTranscriptPreparation.CreateIntelligenceXTranscriptReaderOptions(
                preservesGroupedDefinitionLikeParagraphs: false,
                visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
            var opts = MarkdownRendererPresets.CreateIntelligenceXTranscript();

            Assert.True(opts.ReaderOptions.PreferNarrativeSingleLineDefinitions);
            Assert.Equal(
                expected.InputNormalization!.NormalizeCollapsedOrderedListBoundaries,
                opts.ReaderOptions.InputNormalization!.NormalizeCollapsedOrderedListBoundaries);
            Assert.Equal(
                expected.InputNormalization.NormalizeBrokenTwoLineStrongLeadIns,
                opts.ReaderOptions.InputNormalization.NormalizeBrokenTwoLineStrongLeadIns);
            Assert.Equal(
                expected.InputNormalization.NormalizeHeadingListBoundaries,
                opts.ReaderOptions.InputNormalization.NormalizeHeadingListBoundaries);
            Assert.Equal(
                expected.InputNormalization.NormalizeCompactStrongLabelListBoundaries,
                opts.ReaderOptions.InputNormalization.NormalizeCompactStrongLabelListBoundaries);
            Assert.Equal(
                expected.InputNormalization.NormalizeCompactHeadingBoundaries,
                opts.ReaderOptions.InputNormalization.NormalizeCompactHeadingBoundaries);
            Assert.Equal(
                expected.InputNormalization.NormalizeColonListBoundaries,
                opts.ReaderOptions.InputNormalization.NormalizeColonListBoundaries);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform => transform is MarkdownSimpleDefinitionListParagraphTransform);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform =>
                transform is MarkdownJsonVisualCodeBlockTransform visual
                && visual.LanguageMode == MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscriptMinimal_Reuses_Shared_Transcript_Reader_Contract() {
            var expected = MarkdownTranscriptPreparation.CreateIntelligenceXTranscriptReaderOptions(
                preservesGroupedDefinitionLikeParagraphs: false,
                visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
            var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();

            Assert.True(opts.ReaderOptions.PreferNarrativeSingleLineDefinitions);
            Assert.Equal(
                expected.InputNormalization!.NormalizeCollapsedOrderedListBoundaries,
                opts.ReaderOptions.InputNormalization!.NormalizeCollapsedOrderedListBoundaries);
            Assert.Equal(
                expected.InputNormalization.NormalizeBrokenTwoLineStrongLeadIns,
                opts.ReaderOptions.InputNormalization.NormalizeBrokenTwoLineStrongLeadIns);
            Assert.Equal(
                expected.InputNormalization.NormalizeHeadingListBoundaries,
                opts.ReaderOptions.InputNormalization.NormalizeHeadingListBoundaries);
            Assert.Equal(
                expected.InputNormalization.NormalizeCompactStrongLabelListBoundaries,
                opts.ReaderOptions.InputNormalization.NormalizeCompactStrongLabelListBoundaries);
            Assert.Equal(
                expected.InputNormalization.NormalizeCompactHeadingBoundaries,
                opts.ReaderOptions.InputNormalization.NormalizeCompactHeadingBoundaries);
            Assert.Equal(
                expected.InputNormalization.NormalizeColonListBoundaries,
                opts.ReaderOptions.InputNormalization.NormalizeColonListBoundaries);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform => transform is MarkdownSimpleDefinitionListParagraphTransform);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform =>
                transform is MarkdownJsonVisualCodeBlockTransform visual
                && visual.LanguageMode == MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscriptDesktopShell_BuildsOnMinimalPresetAndEnablesInteractiveVisuals() {
            var minimal = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
            var desktop = MarkdownRendererPresets.CreateIntelligenceXTranscriptDesktopShell();

            Assert.Equal(minimal.HtmlOptions.Style, desktop.HtmlOptions.Style);
            Assert.Equal(minimal.HtmlOptions.CssScopeSelector, desktop.HtmlOptions.CssScopeSelector);
            Assert.Equal(minimal.EnableCodeCopyButtons, desktop.EnableCodeCopyButtons);
            Assert.Equal(minimal.EnableTableCopyButtons, desktop.EnableTableCopyButtons);
            Assert.Equal(minimal.MarkdownPreProcessors.Count, desktop.MarkdownPreProcessors.Count);
            Assert.True(desktop.Mermaid.Enabled);
            Assert.True(desktop.Chart.Enabled);
            Assert.True(desktop.Network.Enabled);
            Assert.False(desktop.Math.Enabled);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateStrict_DoesNotRegister_IntelligenceX_FenceAliases() {
            var opts = MarkdownRendererPresets.CreateStrict();

            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("dataview", StringComparer.OrdinalIgnoreCase));
            Assert.DoesNotContain(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase));
            Assert.DoesNotContain(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-network", StringComparer.OrdinalIgnoreCase));
            Assert.DoesNotContain(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-dataview", StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscript_Registers_IntelligenceX_FenceAliases() {
            var opts = MarkdownRendererPresets.CreateIntelligenceXTranscript();

            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-network", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-dataview", StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownRendererIntelligenceXAdapter_Can_Opt_Generic_Preset_Into_Ix_Aliases() {
            var opts = MarkdownRendererPresets.CreateStrict();

            MarkdownRendererIntelligenceXAdapter.Apply(opts);

            Assert.True(MarkdownRendererIntelligenceXAdapter.IsApplied(opts));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-network", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-dataview", StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownRendererPlugins_IntelligenceXVisuals_Can_Be_Applied_Directly() {
            var opts = MarkdownRendererPresets.CreateStrict();

            opts.ApplyPlugin(MarkdownRendererPlugins.IntelligenceXVisuals);

            Assert.True(opts.HasPlugin(MarkdownRendererPlugins.IntelligenceXVisuals));
            Assert.False(IntelligenceXMarkdownRenderer.HasVisualFenceSchema(opts));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-network", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-dataview", StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownRendererPlugins_IntelligenceXTranscriptVisuals_Can_Be_Applied_Directly() {
            var opts = MarkdownRendererPresets.CreateStrict();

            opts.ApplyPlugin(MarkdownRendererPlugins.IntelligenceXTranscriptVisuals);

            Assert.True(opts.HasPlugin(MarkdownRendererPlugins.IntelligenceXTranscriptVisuals));
            Assert.True(opts.ReaderOptions.PreferNarrativeSingleLineDefinitions);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform => transform is MarkdownSimpleDefinitionListParagraphTransform);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform =>
                transform is MarkdownJsonVisualCodeBlockTransform visual
                && visual.LanguageMode == MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_VisualsPlugin_Carries_VisualFenceSchema() {
            var opts = MarkdownRendererPresets.CreateStrict();

            opts.ApplyPlugin(IntelligenceXMarkdownRenderer.VisualsPlugin);

            Assert.True(opts.HasPlugin(IntelligenceXMarkdownRenderer.VisualsPlugin));
            Assert.True(IntelligenceXMarkdownRenderer.HasVisualFenceSchema(opts));
            Assert.True(opts.TryGetFenceOptionSchema("ix-chart", out var schema));
            Assert.Equal(IntelligenceXMarkdownRenderer.VisualFenceSchema.Id, schema.Id);
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_TranscriptPlugin_Carries_Reader_Contract_And_VisualFenceSchema() {
            var opts = MarkdownRendererPresets.CreateStrict();

            opts.ApplyPlugin(IntelligenceXMarkdownRenderer.TranscriptPlugin);

            Assert.True(opts.HasPlugin(IntelligenceXMarkdownRenderer.TranscriptPlugin));
            Assert.True(IntelligenceXMarkdownRenderer.HasTranscriptContract(opts));
            Assert.True(IntelligenceXMarkdownRenderer.HasVisualFenceSchema(opts));
            Assert.True(opts.ReaderOptions.PreferNarrativeSingleLineDefinitions);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform => transform is MarkdownSimpleDefinitionListParagraphTransform);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform =>
                transform is MarkdownJsonVisualCodeBlockTransform visual
                && visual.LanguageMode == MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_ApplyVisuals_Adds_IxVisualPlugin() {
            var opts = MarkdownRendererPresets.CreateStrict();

            IntelligenceXMarkdownRenderer.ApplyVisuals(opts);

            Assert.True(opts.HasPlugin(IntelligenceXMarkdownRenderer.VisualsPlugin));
            Assert.True(IntelligenceXMarkdownRenderer.HasVisualFenceSchema(opts));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-network", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-dataview", StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_ApplyTranscriptContract_Adds_TranscriptPlugin() {
            var opts = MarkdownRendererPresets.CreateStrict();

            IntelligenceXMarkdownRenderer.ApplyTranscriptContract(opts);

            Assert.True(IntelligenceXMarkdownRenderer.HasTranscriptContract(opts));
            Assert.True(IntelligenceXMarkdownRenderer.HasVisualFenceSchema(opts));
            Assert.True(opts.ReaderOptions.PreferNarrativeSingleLineDefinitions);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform => transform is MarkdownSimpleDefinitionListParagraphTransform);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform =>
                transform is MarkdownJsonVisualCodeBlockTransform visual
                && visual.LanguageMode == MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_ApplyTranscriptContract_Is_Idempotent() {
            var opts = MarkdownRendererPresets.CreateStrict();

            IntelligenceXMarkdownRenderer.ApplyTranscriptContract(opts);
            IntelligenceXMarkdownRenderer.ApplyTranscriptContract(opts);

            Assert.True(IntelligenceXMarkdownRenderer.HasTranscriptContract(opts));
            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase)));
            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-network", StringComparer.OrdinalIgnoreCase)));
            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-dataview", StringComparer.OrdinalIgnoreCase)));
            Assert.Equal(1, opts.ReaderOptions.DocumentTransforms.Count(transform => transform is MarkdownSimpleDefinitionListParagraphTransform));
            Assert.Equal(1, opts.ReaderOptions.DocumentTransforms.Count(transform =>
                transform is MarkdownJsonVisualCodeBlockTransform visual
                && visual.LanguageMode == MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence));
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_ApplyTranscriptCompatibility_Matches_Core_Composition() {
            var viaPackage = MarkdownRendererPresets.CreateStrict();
            var viaCore = MarkdownRendererPresets.CreateStrict();

            IntelligenceXMarkdownRenderer.ApplyTranscriptCompatibility(viaPackage);
            viaCore.ApplyFeaturePack(MarkdownRendererFeaturePacks.IntelligenceXTranscriptCompatibility);

            Assert.Equal(viaCore.MarkdownPreProcessors.Count, viaPackage.MarkdownPreProcessors.Count);
            Assert.True(viaPackage.HasFeaturePack(MarkdownRendererFeaturePacks.IntelligenceXTranscriptCompatibility));
            Assert.True(IntelligenceXMarkdownRenderer.HasTranscriptCompatibility(viaPackage));
            Assert.True(IntelligenceXMarkdownRenderer.HasTranscriptContract(viaPackage));
            Assert.True(IntelligenceXMarkdownRenderer.HasVisualFenceSchema(viaPackage));
            Assert.Equal(viaCore.ReaderOptions.PreferNarrativeSingleLineDefinitions, viaPackage.ReaderOptions.PreferNarrativeSingleLineDefinitions);
            Assert.Equal(viaCore.ReaderOptions.DocumentTransforms.Count, viaPackage.ReaderOptions.DocumentTransforms.Count);
            Assert.Equal(
                viaCore.FencedCodeBlockRenderers.SelectMany(renderer => renderer.Languages).OrderBy(value => value, StringComparer.OrdinalIgnoreCase),
                viaPackage.FencedCodeBlockRenderers.SelectMany(renderer => renderer.Languages).OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_TranscriptCompatibilityPack_Carries_VisualFenceSchema() {
            var opts = MarkdownRendererPresets.CreateStrict();

            opts.ApplyFeaturePack(IntelligenceXMarkdownRenderer.TranscriptCompatibilityPack);

            Assert.True(opts.HasFeaturePack(IntelligenceXMarkdownRenderer.TranscriptCompatibilityPack));
            Assert.True(IntelligenceXMarkdownRenderer.HasTranscriptCompatibility(opts));
            Assert.True(IntelligenceXMarkdownRenderer.HasTranscriptContract(opts));
            Assert.True(IntelligenceXMarkdownRenderer.HasVisualFenceSchema(opts));
            Assert.True(opts.ReaderOptions.PreferNarrativeSingleLineDefinitions);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform => transform is MarkdownSimpleDefinitionListParagraphTransform);
            Assert.Contains(opts.ReaderOptions.DocumentTransforms, transform =>
                transform is MarkdownJsonVisualCodeBlockTransform visual
                && visual.LanguageMode == MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_VisualFenceSchema_Is_Resolvable_By_Renderer_Options() {
            var opts = MarkdownRendererPresets.CreateStrict();

            IntelligenceXMarkdownRenderer.ApplyVisualFenceSchema(opts);

            Assert.True(IntelligenceXMarkdownRenderer.HasVisualFenceSchema(opts));
            Assert.True(opts.TryGetFenceOptionSchema("ix-chart", out var chartSchema));
            Assert.Equal(IntelligenceXMarkdownRenderer.VisualFenceSchema.Id, chartSchema.Id);
            Assert.True(opts.TryGetFenceOptionSchema("ix-network", out var networkSchema));
            Assert.Equal(IntelligenceXMarkdownRenderer.VisualFenceSchema.Id, networkSchema.Id);
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_Can_Parse_Typed_Visual_Fence_Options() {
            var parsed = IntelligenceXMarkdownRenderer.ParseVisualFenceOptions(
                "ix-chart {#quarterly-summary .wide .accent title=\"Quarterly Revenue\" pinned theme=\"amber\" variant=compact view=timeline maxItems=12}");

            Assert.Equal("ix-chart", parsed.Language);
            Assert.Equal("ix-chart {#quarterly-summary .wide .accent title=\"Quarterly Revenue\" pinned theme=\"amber\" variant=compact view=timeline maxItems=12}", parsed.InfoString);
            Assert.Equal("quarterly-summary", parsed.ElementId);
            Assert.Equal(new[] { "wide", "accent" }, parsed.Classes);
            Assert.True(parsed.HasClass("wide"));
            Assert.Equal("Quarterly Revenue", parsed.Title);
            Assert.True(parsed.Pinned);
            Assert.Equal("amber", parsed.Theme);
            Assert.Equal("compact", parsed.Variant);
            Assert.Equal("timeline", parsed.View);
            Assert.Equal(12, parsed.MaxItems);
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_Can_Parse_Typed_Visual_Fence_Options_From_Shared_Fence_Info() {
            var fenceInfo = MarkdownCodeFenceInfo.Parse("ix-network title=\"Relationship Map\" pin mode=graph limit=8");
            var parsed = IntelligenceXMarkdownRenderer.ParseVisualFenceOptions(fenceInfo);

            Assert.Equal("ix-network", parsed.Language);
            Assert.Equal("Relationship Map", parsed.Title);
            Assert.True(parsed.Pinned);
            Assert.Equal("graph", parsed.View);
            Assert.Equal(8, parsed.MaxItems);
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_VisualFenceSchema_Parses_And_Validates_Registered_Options() {
            var opts = MarkdownRendererPresets.CreateStrict();
            IntelligenceXMarkdownRenderer.ApplyVisualFenceSchema(opts);
            var fenceInfo = MarkdownCodeFenceInfo.Parse("ix-chart title=\"Quarterly Revenue\" pin palette=amber style=compact mode=timeline limit=0 custom=true");

            Assert.True(opts.TryParseFenceOptions("ix-chart", fenceInfo, out var parsed));
            Assert.False(parsed.IsValid);
            Assert.True(parsed.TryGetBoolean("pinned", out var pinned));
            Assert.True(pinned);
            Assert.True(parsed.TryGetString("theme", out var theme));
            Assert.Equal("amber", theme);
            Assert.True(parsed.TryGetString("variant", out var variant));
            Assert.Equal("compact", variant);
            Assert.True(parsed.TryGetString("view", out var view));
            Assert.Equal("timeline", view);
            Assert.Contains("maxItems", parsed.Errors.Keys, StringComparer.OrdinalIgnoreCase);
            Assert.Contains("custom", parsed.UnknownOptions);
            Assert.DoesNotContain("title", parsed.UnknownOptions);
        }

        [Fact]
        public void MarkdownRendererFeaturePacks_IntelligenceXTranscriptCompatibility_Is_Idempotent_And_Tracked() {
            var opts = MarkdownRendererPresets.CreateStrict();

            opts.ApplyFeaturePack(MarkdownRendererFeaturePacks.IntelligenceXTranscriptCompatibility);
            opts.ApplyFeaturePack(MarkdownRendererFeaturePacks.IntelligenceXTranscriptCompatibility);

            Assert.True(opts.HasFeaturePack(MarkdownRendererFeaturePacks.IntelligenceXTranscriptCompatibility));
            Assert.Contains(opts.AppliedFeaturePackIds, id => string.Equals(id, "officeimo.intelligencex.transcript-compatibility", StringComparison.OrdinalIgnoreCase));
            Assert.Equal(1, opts.AppliedFeaturePackIds.Count(id => string.Equals(id, "officeimo.intelligencex.transcript-compatibility", StringComparison.OrdinalIgnoreCase)));
            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase)));
            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-network", StringComparer.OrdinalIgnoreCase)));
            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-dataview", StringComparer.OrdinalIgnoreCase)));
            Assert.Equal(2, opts.MarkdownPreProcessors.Count);
            Assert.Equal(1, opts.ReaderOptions.DocumentTransforms.Count(transform => transform is MarkdownSimpleDefinitionListParagraphTransform));
            Assert.Equal(1, opts.ReaderOptions.DocumentTransforms.Count(transform =>
                transform is MarkdownJsonVisualCodeBlockTransform visual
                && visual.LanguageMode == MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence));
        }

        [Fact]
        public void IntelligenceXTranscriptCompatibilityPack_Can_Upgrade_Legacy_Json_Visuals_On_Generic_Strict_Renderer() {
            const string markdown = """
```json
{"type":"bar","data":{"labels":["A"],"datasets":[{"label":"Count","data":[1]}]}}
```
""";
            var opts = MarkdownRendererPresets.CreateStrict();
            opts.Chart.Enabled = true;

            opts.ApplyFeaturePack(IntelligenceXMarkdownRenderer.TranscriptCompatibilityPack);
            var html = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, opts);

            Assert.Contains("class=\"omd-visual omd-chart\"", html, StringComparison.Ordinal);
            Assert.Contains("data-omd-fence-language=\"ix-chart\"", html, StringComparison.Ordinal);
        }

        [Fact]
        public void IntelligenceXMarkdownRenderer_CreateTranscriptDesktopShell_Matches_CorePreset() {
            var viaPackage = IntelligenceXMarkdownRenderer.CreateTranscriptDesktopShell();
            var viaCore = MarkdownRendererPresets.CreateIntelligenceXTranscriptDesktopShell();

            Assert.Equal(viaCore.HtmlOptions.Style, viaPackage.HtmlOptions.Style);
            Assert.Equal(viaCore.HtmlOptions.CssScopeSelector, viaPackage.HtmlOptions.CssScopeSelector);
            Assert.Equal(viaCore.EnableCodeCopyButtons, viaPackage.EnableCodeCopyButtons);
            Assert.Equal(viaCore.EnableTableCopyButtons, viaPackage.EnableTableCopyButtons);
            Assert.Equal(viaCore.Mermaid.Enabled, viaPackage.Mermaid.Enabled);
            Assert.Equal(viaCore.Chart.Enabled, viaPackage.Chart.Enabled);
            Assert.Equal(viaCore.Network.Enabled, viaPackage.Network.Enabled);
            Assert.Equal(viaCore.Math.Enabled, viaPackage.Math.Enabled);
            Assert.True(IntelligenceXMarkdownRenderer.HasVisualFenceSchema(viaPackage));
            Assert.Equal(
                viaCore.FencedCodeBlockRenderers.SelectMany(renderer => renderer.Languages).OrderBy(value => value, StringComparer.OrdinalIgnoreCase),
                viaPackage.FencedCodeBlockRenderers.SelectMany(renderer => renderer.Languages).OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownRendererOptions_Defaults_Install_GenericVisualPlugin() {
            var opts = new MarkdownRendererOptions();

            Assert.True(opts.HasPlugin(MarkdownRendererPlugins.GenericVisuals));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("chart", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("network", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("dataview", StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownRendererIntelligenceXAdapter_Is_Idempotent() {
            var opts = MarkdownRendererPresets.CreateStrict();

            MarkdownRendererIntelligenceXAdapter.Apply(opts);
            MarkdownRendererIntelligenceXAdapter.Apply(opts);

            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase)));
            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-network", StringComparer.OrdinalIgnoreCase)));
            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-dataview", StringComparer.OrdinalIgnoreCase)));
        }

        [Fact]
        public void MarkdownRendererIntelligenceXLegacyMigration_AddsLegacyHeadingCleanupPreProcessor_OnlyOnce() {
            var opts = MarkdownRendererPresets.CreateStrict();

            MarkdownRendererIntelligenceXLegacyMigration.Apply(opts);
            int once = opts.MarkdownPreProcessors.Count;
            MarkdownRendererIntelligenceXLegacyMigration.Apply(opts);

            Assert.Equal(once, opts.MarkdownPreProcessors.Count);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscriptMinimal_RepairsLegacyToolHeadingArtifacts() {
            var markdown = """
[Cached evidence fallback]

Recent evidence:
- eventlog_top_events: ### Top 30 recent events (preview)

#### ad_environment_discover
### Active Directory: Environment Discovery
""";

            var strict = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, MarkdownRendererPresets.CreateStrictMinimal());
            var chat = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal());

            Assert.Contains("eventlog_top_events", strict, StringComparison.Ordinal);
            Assert.Contains("Top 30 recent events (preview)", chat, StringComparison.Ordinal);
            Assert.DoesNotContain("eventlog_top_events:", chat, StringComparison.Ordinal);
            Assert.DoesNotContain("ad_environment_discover", chat, StringComparison.Ordinal);
            Assert.Contains("Active Directory: Environment Discovery", chat, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscriptMinimal_RepairsHostLabelBulletsAndBrokenResultLeadIns() {
            var markdown = """
-AD1
healthy for directory access

**Result
all 5 are healthy for directory access** with recommended LDAPS endpoints.
""";

            var strict = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, MarkdownRendererPresets.CreateStrictMinimal());
            var chat = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal());

            Assert.Contains("AD1", strict, StringComparison.Ordinal);
            Assert.Contains("Result", strict, StringComparison.Ordinal);
            Assert.Contains("AD1 healthy for directory access", chat, StringComparison.Ordinal);
            Assert.Contains("Result:", chat, StringComparison.Ordinal);
            Assert.DoesNotContain("<strong>Result\n", chat, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscriptMinimal_Flattens_Grouped_Simple_Definitions() {
            var markdown = """
Status: healthy
Impact: none
""";

            var chat = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal());

            Assert.DoesNotContain("<dl>", chat, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<p>Status: healthy</p>", chat, StringComparison.Ordinal);
            Assert.Contains("<p>Impact: none</p>", chat, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscript_RepairsCachedEvidenceNetworkTransportArtifacts() {
            var markdown = """
ix:cached-tool-evidence:v1

Recent scope graph:

```json
{
  "nodes": [
    { "id": "forest_ad.evotec.xyz", "label": "Forest: ad.evotec.xyz" }
  ],
  "edges": [
    { "source": "forest_ad.evotec.xyz", "target": "domain_ad.evotec.xyz", "label": "contains" }
  ]
}
```

Indented fallback:

    {
      "nodes": [
        { "id": "domain_ad.evotec.xyz", "label": "Domain: ad.evotec.xyz" }
      ],
      "edges": [
        { "source": "domain_ad.evotec.xyz", "target": "dc_ad0.ad.evotec.xyz", "label": "hosts" }
      ]
    }
""";

            var strictOptions = MarkdownRendererPresets.CreateStrict();
            strictOptions.Network.Enabled = true;
            var strict = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strictOptions);

            var chatOptions = MarkdownRendererPresets.CreateIntelligenceXTranscript();
            chatOptions.Network.Enabled = true;
            var chat = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chatOptions);

            Assert.Contains("cached-tool-evidence", strict, StringComparison.Ordinal);
            Assert.Contains("language-json", strict, StringComparison.Ordinal);
            Assert.DoesNotContain("cached-tool-evidence", chat, StringComparison.Ordinal);
            Assert.Contains("data-omd-fence-language=\"ix-network\"", chat, StringComparison.Ordinal);
            Assert.Equal(2, CountOccurrences(chat, "data-omd-fence-language=\"ix-network\""));
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscript_UpgradesPlainLegacyJsonNetworkFence() {
            var markdown = """
```json
{
  "nodes": [
    { "id": "forest_ad.evotec.xyz", "label": "Forest: ad.evotec.xyz" }
  ],
  "edges": [
    { "source": "forest_ad.evotec.xyz", "target": "domain_ad.evotec.xyz", "label": "contains" }
  ]
}
```
""";

            var strictOptions = MarkdownRendererPresets.CreateStrict();
            strictOptions.Network.Enabled = true;
            var strict = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strictOptions);

            var transcriptOptions = MarkdownRendererPresets.CreateIntelligenceXTranscript();
            transcriptOptions.Network.Enabled = true;
            var transcript = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, transcriptOptions);

            Assert.Contains("language-json", strict, StringComparison.Ordinal);
            Assert.Contains("data-omd-fence-language=\"ix-network\"", transcript, StringComparison.Ordinal);
            Assert.DoesNotContain("language-json", transcript, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateIntelligenceXTranscript_RepairsCachedEvidenceChartAndDataViewTransportArtifacts() {
            var markdown = """
ix:cached-tool-evidence:v1

Chart preview:

```json
{
  "type": "bar",
  "data": {
    "labels": [ "A" ],
    "datasets": [
      { "label": "Count", "data": [ 1 ] }
    ]
  }
}
```

Dataview preview:

```json
{
  "title": "Replication Summary",
  "summary": "Latest replication posture",
  "kind": "ix_tool_dataview_v1",
  "call_id": "call_123",
  "rows": [
    [ "Server", "Fails" ],
    [ "AD0", "0" ],
    [ "AD1", "1" ]
  ]
}
```
""";

            var strictOptions = MarkdownRendererPresets.CreateStrict();
            strictOptions.Chart.Enabled = true;
            var strict = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strictOptions);

            var chatOptions = MarkdownRendererPresets.CreateIntelligenceXTranscript();
            chatOptions.Chart.Enabled = true;
            var chat = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chatOptions);

            Assert.Contains("cached-tool-evidence", strict, StringComparison.Ordinal);
            Assert.Equal(2, CountOccurrences(strict, "language-json"));
            Assert.DoesNotContain("cached-tool-evidence", chat, StringComparison.Ordinal);
            Assert.Contains("class=\"omd-visual omd-chart\"", chat, StringComparison.Ordinal);
            Assert.Contains("data-omd-fence-language=\"ix-chart\"", chat, StringComparison.Ordinal);
            Assert.Contains("class=\"omd-visual omd-dataview\"", chat, StringComparison.Ordinal);
            Assert.Contains("data-omd-fence-language=\"ix-dataview\"", chat, StringComparison.Ordinal);
        }

        private static int CountOccurrences(string text, string value) {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(value)) {
                return 0;
            }

            var count = 0;
            var index = 0;
            while (true) {
                index = text.IndexOf(value, index, StringComparison.Ordinal);
                if (index < 0) {
                    return count;
                }

                count++;
                index += value.Length;
            }
        }

        [Fact]
        public void MarkdownRenderer_RenderUpdateScript_ProducesUpdateContentCall() {
            var opts = MarkdownRendererPresets.CreateIntelligenceXTranscript();
            var js = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderUpdateScript("**bold**", opts);

            Assert.StartsWith("updateContent(", js);
            Assert.EndsWith(");", js);
            Assert.Contains("markdown-body", js);
        }

        [Fact]
        public void HtmlStyle_ChatAuto_Includes_Bubble_Css_Classes() {
            var doc = MarkdownReader.Parse("Hello");
            var parts = doc.ToHtmlParts(new HtmlOptions { Kind = HtmlKind.Fragment, Style = HtmlStyle.ChatAuto });

            Assert.Contains(".omd-chat-bubble", parts.Css, StringComparison.Ordinal);
            Assert.Contains(".omd-chat-row", parts.Css, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownRenderer_Can_Wrap_As_ChatBubble() {
            var opts = MarkdownRendererPresets.CreateIntelligenceXTranscript();
            var bubble = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderChatBubbleBodyHtml("Hello", ChatMessageRole.User, opts);

            Assert.Contains("omd-chat-row", bubble, StringComparison.Ordinal);
            Assert.Contains("omd-chat-bubble", bubble, StringComparison.Ordinal);
            Assert.Contains("omd-role-user", bubble, StringComparison.Ordinal);
            Assert.Contains("markdown-body", bubble, StringComparison.Ordinal);
        }
    }
}
