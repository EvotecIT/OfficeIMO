using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_Footnotes() {
            string md = "Text with footnote[^1].\n\n[^1]: Footnote text";
            using var doc = md.LoadFromMarkdown();
            Assert.NotNull(doc.FootNotes);
            Assert.Single(doc.FootNotes);
            Assert.Equal("Footnote text", doc.FootNotes![0].Paragraphs[1].Text);
        }
    }
}
