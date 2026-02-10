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
