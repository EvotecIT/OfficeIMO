using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlBlockquoteCitations {
        [Fact]
        public void BlockquoteWithCitationCreatesFootnote() {
            string html = "<blockquote cite=\"https://example.com\"><p>First</p><p>Second</p></blockquote>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Equal("First", doc.Paragraphs[0].Text);
            Assert.Contains(doc.Paragraphs, p => p.Text == "Second");
            var footNotes = doc.FootNotes;
            Assert.NotNull(footNotes);
            Assert.Single(footNotes!);
            Assert.Equal("https://example.com", footNotes![0].Paragraphs![1].Text);
        }
    }
}