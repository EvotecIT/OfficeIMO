using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Inline_Nesting_Tests {
        [Fact]
        public void Parses_Triple_Star_As_BoldItalic() {
            var doc = MarkdownReader.Parse("***Both***");
            var outMd = doc.ToMarkdown().Trim();
            Assert.Equal("***Both***", outMd);
            var html = doc.ToHtml().Trim();
            Assert.Contains("<strong><em>Both</em></strong>", html);
        }

        [Fact]
        public void Parses_Link_With_Parentheses() {
            var md = "See [site](https://example.com/foo(bar)/baz).";
            var doc = MarkdownReader.Parse(md);
            var outMd = doc.ToMarkdown();
            Assert.Contains("(https://example.com/foo(bar)/baz)", outMd);
        }
    }
}

