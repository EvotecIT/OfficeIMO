using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_ClassStyles_Paragraph() {
            string html = "<p class=\"title\">Text</p>";
            var options = new HtmlToWordOptions();
            options.ClassStyles["title"] = WordParagraphStyles.Heading1;
            var doc = html.LoadFromHtml(options);
            Assert.Equal(WordParagraphStyles.Heading1, doc.Paragraphs[0].Style);
        }

        [Fact]
        public void HtmlToWord_ClassStyles_ListItem() {
            string html = "<ul><li class=\"special\">Item</li></ul>";
            var options = new HtmlToWordOptions();
            options.ClassStyles["special"] = WordParagraphStyles.Heading2;
            var doc = html.LoadFromHtml(options);
            Assert.Equal(WordParagraphStyles.Heading2, doc.Paragraphs[0].Style);
        }
    }
}
