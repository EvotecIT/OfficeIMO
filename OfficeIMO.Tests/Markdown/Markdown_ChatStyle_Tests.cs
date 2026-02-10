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
    }
}

