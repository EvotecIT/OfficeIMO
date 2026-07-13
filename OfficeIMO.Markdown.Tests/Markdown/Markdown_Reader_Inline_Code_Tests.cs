using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Inline_Code_Tests {
        [Fact]
        public void Double_Backticks_With_Inner_Backticks() {
            var md = "``Use `backticks` inside``";
            var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md);
            var html = doc.ToHtmlFragment();
            Assert.Contains("Use `backticks` inside", html);
        }

        [Fact]
        public void Triple_Backticks_Code_Span() {
            var md = "start ```a ` b``` end";
            var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md);
            var html = doc.ToHtmlFragment();
            Assert.Contains("a ` b", html);
        }

        [Fact]
        public void Unmatched_MultiBacktick_Run_Remains_Literal() {
            var md = "``` c`sharp";
            var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md);
            var html = doc.ToHtmlFragment();

            Assert.Contains("``` c`sharp", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<code>c</code>", html, StringComparison.Ordinal);
        }
    }
}
