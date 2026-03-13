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
            Assert.False(opts.ReaderOptions.AutolinkUrls);
            Assert.False(opts.ReaderOptions.AutolinkWwwUrls);
            Assert.False(opts.ReaderOptions.AutolinkEmails);
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
            MarkdownRendererPresets.ApplyChatPresentation(composed, enableCopyButtons: false);
            MarkdownRendererIntelligenceXAdapter.Apply(composed);

            var wrapper = MarkdownRendererPresets.CreateChatStrictMinimal();

            Assert.Equal(wrapper.HtmlOptions.Style, composed.HtmlOptions.Style);
            Assert.Equal(wrapper.HtmlOptions.CssScopeSelector, composed.HtmlOptions.CssScopeSelector);
            Assert.Equal(wrapper.EnableCodeCopyButtons, composed.EnableCodeCopyButtons);
            Assert.Equal(wrapper.EnableTableCopyButtons, composed.EnableTableCopyButtons);
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
