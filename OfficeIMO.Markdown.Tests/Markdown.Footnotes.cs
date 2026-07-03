using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_Footnotes() {
            string md = "Text with footnote[^1].\n\n[^1]: Footnote text";
            using var doc = md.LoadFromMarkdown();
            var footNotes = doc.FootNotes;
            Assert.NotNull(footNotes);
            Assert.Single(footNotes!);
            Assert.Equal("Footnote text", footNotes![0].Paragraphs![1].Text);
        }
    }
}
