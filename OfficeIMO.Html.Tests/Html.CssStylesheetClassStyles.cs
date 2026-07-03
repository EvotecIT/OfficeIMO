using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_CssStylesheet_Paragraph() {
            string html = "<style>.title { font-weight:bold; font-size:32px; }</style><p class=\"title\">Text</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            Assert.Equal(WordParagraphStyles.Heading1, doc.Paragraphs[0].Style);
        }

        [Fact]
        public void HtmlToWord_CssStylesheet_ListItem() {
            string html = "<style>.special { font-weight:bold; font-size:24px; }</style><ul><li class=\"special\">Item</li></ul>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            Assert.Equal(WordParagraphStyles.Heading2, doc.Paragraphs[0].Style);
        }
    }
}

