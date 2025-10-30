using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void Markdown_DuplicateHeadingsProduceUniqueAnchors() {
            var doc = MarkdownDoc.Create()
                .Toc(opts => { opts.IncludeTitle = false; opts.MinLevel = 2; opts.MaxLevel = 3; }, placeAtTop: true)
                .H1("Intro")
                .H2("Repeat me")
                .H2("Repeat me")
                .H2("Repeat me");

            string markdown = doc.ToMarkdown();
            Assert.Contains("- [Repeat me](#repeat-me)", markdown);
            Assert.Contains("- [Repeat me](#repeat-me-1)", markdown);
            Assert.Contains("- [Repeat me](#repeat-me-2)", markdown);

            string html = doc.ToHtmlFragment();
            Assert.Contains("<h2 id=\"repeat-me\"", html);
            Assert.Contains("<h2 id=\"repeat-me-1\"", html);
            Assert.Contains("<h2 id=\"repeat-me-2\"", html);
            Assert.Contains("href=\"#repeat-me\"", html);
            Assert.Contains("href=\"#repeat-me-1\"", html);
            Assert.Contains("href=\"#repeat-me-2\"", html);
        }
    }
}
