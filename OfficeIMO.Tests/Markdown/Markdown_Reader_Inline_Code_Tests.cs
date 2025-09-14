using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Inline_Code_Tests {
        [Fact]
        public void Double_Backticks_With_Inner_Backticks() {
            var md = "``Use `backticks` inside``";
            var doc = MarkdownReader.Parse(md);
            var html = doc.ToHtml();
            Assert.Contains("Use `backticks` inside", html);
        }

        [Fact]
        public void Triple_Backticks_Code_Span() {
            var md = "start ```a ` b``` end";
            var doc = MarkdownReader.Parse(md);
            var html = doc.ToHtml();
            Assert.Contains("a ` b", html);
        }
    }
}
