using System;
using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests {
    public class Markdown_ChatStyle_Tests {
        [Fact]
        public void HtmlStyle_ChatAuto_EmitsChatMarkerAndAutoThemeCss() {
            var doc = MarkdownReader.Parse("Hello");
            var parts = doc.ToHtmlParts(new HtmlOptions { Kind = HtmlKind.Fragment, Style = HtmlStyle.ChatAuto });

            Assert.Contains("omd-chat", parts.Css);
            Assert.Contains("prefers-color-scheme", parts.Css);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateChatStrict_UsesChatStyleAndScopedCss() {
            var opts = MarkdownRendererPresets.CreateChatStrict();
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
        public void MarkdownRendererPresets_CreateChatStrictPortable_Disables_OfficeImoOnly_Reader_Extensions() {
            var opts = MarkdownRendererPresets.CreateChatStrictPortable();

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
        public void MarkdownRendererPresets_CreateChatStrict_Can_Target_Gfm_Profile() {
            var opts = MarkdownRendererPresets.CreateChatStrict(MarkdownReaderOptions.MarkdownDialectProfile.GitHubFlavoredMarkdown);

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
        public void MarkdownRendererPresets_CreateChatStrict_BuildsOnStrictDefaults() {
            var strict = MarkdownRendererPresets.CreateStrict();
            var chat = MarkdownRendererPresets.CreateChatStrict();

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

            Assert.Equal(HtmlStyle.ChatAuto, opts.HtmlOptions.Style);
            Assert.Equal("#omdRoot article.markdown-body", opts.HtmlOptions.CssScopeSelector);
            Assert.False(opts.EnableCodeCopyButtons);
            Assert.False(opts.EnableTableCopyButtons);
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-network", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(opts.FencedCodeBlockRenderers, renderer => renderer.Languages.Contains("ix-dataview", StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownRendererPresets_CreateChatStrictMinimal_Matches_Composed_Generic_Preset() {
            var composed = MarkdownRendererPresets.CreateStrictMinimal();
            composed.NormalizeWrappedSignalFlowStrongRuns = true;
            composed.NormalizeCollapsedMetricChains = true;
            composed.NormalizeHostLabelBulletArtifacts = true;
            composed.NormalizeStandaloneHashHeadingSeparators = true;
            composed.NormalizeBrokenTwoLineStrongLeadIns = true;
            composed.NormalizeDanglingTrailingStrongListClosers = true;
            composed.NormalizeMetricValueStrongRuns = true;
            MarkdownRendererPresets.ApplyChatPresentation(composed, enableCopyButtons: false);
            MarkdownRendererIntelligenceXAdapter.Apply(composed);

            var wrapper = MarkdownRendererPresets.CreateChatStrictMinimal();

            Assert.Equal(wrapper.HtmlOptions.Style, composed.HtmlOptions.Style);
            Assert.Equal(wrapper.HtmlOptions.CssScopeSelector, composed.HtmlOptions.CssScopeSelector);
            Assert.Equal(wrapper.EnableCodeCopyButtons, composed.EnableCodeCopyButtons);
            Assert.Equal(wrapper.EnableTableCopyButtons, composed.EnableTableCopyButtons);
            Assert.Equal(wrapper.NormalizeWrappedSignalFlowStrongRuns, composed.NormalizeWrappedSignalFlowStrongRuns);
            Assert.Equal(wrapper.NormalizeCollapsedMetricChains, composed.NormalizeCollapsedMetricChains);
            Assert.Equal(wrapper.NormalizeHostLabelBulletArtifacts, composed.NormalizeHostLabelBulletArtifacts);
            Assert.Equal(wrapper.NormalizeStandaloneHashHeadingSeparators, composed.NormalizeStandaloneHashHeadingSeparators);
            Assert.Equal(wrapper.NormalizeBrokenTwoLineStrongLeadIns, composed.NormalizeBrokenTwoLineStrongLeadIns);
            Assert.Equal(wrapper.NormalizeDanglingTrailingStrongListClosers, composed.NormalizeDanglingTrailingStrongListClosers);
            Assert.Equal(wrapper.NormalizeMetricValueStrongRuns, composed.NormalizeMetricValueStrongRuns);
            Assert.Equal(wrapper.MarkdownPreProcessors.Count, composed.MarkdownPreProcessors.Count);
            Assert.Equal(
                wrapper.FencedCodeBlockRenderers.SelectMany(renderer => renderer.Languages).OrderBy(value => value, StringComparer.OrdinalIgnoreCase),
                composed.FencedCodeBlockRenderers.SelectMany(renderer => renderer.Languages).OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
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
        public void MarkdownRendererPresets_CreateChatStrict_Registers_IntelligenceX_FenceAliases() {
            var opts = MarkdownRendererPresets.CreateChatStrict();

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
        public void MarkdownRendererIntelligenceXAdapter_Is_Idempotent() {
            var opts = MarkdownRendererPresets.CreateStrict();

            MarkdownRendererIntelligenceXAdapter.Apply(opts);
            MarkdownRendererIntelligenceXAdapter.Apply(opts);

            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-chart", StringComparer.OrdinalIgnoreCase)));
            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-network", StringComparer.OrdinalIgnoreCase)));
            Assert.Equal(1, opts.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("ix-dataview", StringComparer.OrdinalIgnoreCase)));
        }

        [Fact]
        public void MarkdownRendererIntelligenceXAdapter_AddsLegacyHeadingCleanupPreProcessor_OnlyOnce() {
            var opts = MarkdownRendererPresets.CreateStrict();

            MarkdownRendererIntelligenceXAdapter.Apply(opts);
            int once = opts.MarkdownPreProcessors.Count;
            MarkdownRendererIntelligenceXAdapter.Apply(opts);

            Assert.Equal(once, opts.MarkdownPreProcessors.Count);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateChatStrictMinimal_RepairsLegacyToolHeadingArtifacts() {
            var markdown = """
[Cached evidence fallback]

Recent evidence:
- eventlog_top_events: ### Top 30 recent events (preview)

#### ad_environment_discover
### Active Directory: Environment Discovery
""";

            var strict = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, MarkdownRendererPresets.CreateStrictMinimal());
            var chat = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, MarkdownRendererPresets.CreateChatStrictMinimal());

            Assert.Contains("eventlog_top_events", strict, StringComparison.Ordinal);
            Assert.Contains("Top 30 recent events (preview)", chat, StringComparison.Ordinal);
            Assert.DoesNotContain("eventlog_top_events:", chat, StringComparison.Ordinal);
            Assert.DoesNotContain("ad_environment_discover", chat, StringComparison.Ordinal);
            Assert.Contains("Active Directory: Environment Discovery", chat, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateChatStrictMinimal_RepairsHostLabelBulletsAndBrokenResultLeadIns() {
            var markdown = """
-AD1
healthy for directory access

**Result
all 5 are healthy for directory access** with recommended LDAPS endpoints.
""";

            var strict = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, MarkdownRendererPresets.CreateStrictMinimal());
            var chat = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, MarkdownRendererPresets.CreateChatStrictMinimal());

            Assert.Contains("AD1", strict, StringComparison.Ordinal);
            Assert.Contains("Result", strict, StringComparison.Ordinal);
            Assert.Contains("AD1 healthy for directory access", chat, StringComparison.Ordinal);
            Assert.Contains("Result:", chat, StringComparison.Ordinal);
            Assert.DoesNotContain("<strong>Result\n", chat, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownRendererPresets_CreateChatStrict_RepairsCachedEvidenceNetworkTransportArtifacts() {
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

            var chatOptions = MarkdownRendererPresets.CreateChatStrict();
            chatOptions.Network.Enabled = true;
            var chat = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chatOptions);

            Assert.Contains("cached-tool-evidence", strict, StringComparison.Ordinal);
            Assert.Contains("language-json", strict, StringComparison.Ordinal);
            Assert.DoesNotContain("cached-tool-evidence", chat, StringComparison.Ordinal);
            Assert.Contains("data-omd-fence-language=\"ix-network\"", chat, StringComparison.Ordinal);
            Assert.Equal(2, CountOccurrences(chat, "data-omd-fence-language=\"ix-network\""));
        }

        [Fact]
        public void MarkdownRendererPresets_CreateChatStrict_RepairsCachedEvidenceChartAndDataViewTransportArtifacts() {
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

            var chatOptions = MarkdownRendererPresets.CreateChatStrict();
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
            var opts = MarkdownRendererPresets.CreateChatStrict();
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
            var opts = MarkdownRendererPresets.CreateChatStrict();
            var bubble = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderChatBubbleBodyHtml("Hello", ChatMessageRole.User, opts);

            Assert.Contains("omd-chat-row", bubble, StringComparison.Ordinal);
            Assert.Contains("omd-chat-bubble", bubble, StringComparison.Ordinal);
            Assert.Contains("omd-role-user", bubble, StringComparison.Ordinal);
            Assert.Contains("markdown-body", bubble, StringComparison.Ordinal);
        }
    }
}
