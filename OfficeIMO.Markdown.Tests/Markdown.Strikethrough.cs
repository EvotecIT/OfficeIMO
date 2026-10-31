using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void Markdown_Strikethrough_RoundTrip() {
            string md = "This is ~~strike~~ text";

            var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions());
            var paragraph = doc.Paragraphs[0];
            var run = paragraph.GetRuns().First(r => r.Strike);

            Assert.Equal("strike", run.Text);

            string roundTrip = doc.ToMarkdown(new WordMarkdownExportOptions()).ToMarkdown();
            Assert.Contains("~~strike~~", roundTrip);
        }
    }
}

