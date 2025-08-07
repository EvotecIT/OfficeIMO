using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void Markdown_Strikethrough_RoundTrip() {
            string md = "This is ~~strike~~ text";

            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());
            var paragraph = doc.Paragraphs[0];
            var run = paragraph.GetRuns().First(r => r.Strike);

            Assert.Equal("strike", run.Text);

            string roundTrip = doc.ToMarkdown(new WordToMarkdownOptions());
            Assert.Contains("~~strike~~", roundTrip);
        }
    }
}

