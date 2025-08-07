using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_RoundtripFootNotes() {
            string markdown = "Paragraph one[^1] and two[^2].\n\nAnother paragraph[^3].\n\n[^1]: First footnote\n[^2]: Second footnote\n[^3]: Third footnote\n";
            var doc = markdown.LoadFromMarkdown(new MarkdownToWordOptions());
            Assert.Equal(3, doc.FootNotes.Count);
            Assert.Contains("First footnote", doc.FootNotes[0].Paragraphs[1].Text);
            Assert.Contains("Second footnote", doc.FootNotes[1].Paragraphs[1].Text);
            Assert.Contains("Third footnote", doc.FootNotes[2].Paragraphs[1].Text);
        }
    }
}
