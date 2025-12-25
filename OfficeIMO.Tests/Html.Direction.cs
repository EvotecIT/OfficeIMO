using System.Linq;

using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_ParagraphDirection_DirAttribute() {
            string html = "<div dir=\"rtl\"><p>Alpha</p></div>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var paragraph = doc.Paragraphs.First(p => p.Text.Contains("Alpha"));
            Assert.True(paragraph.BiDi);
        }

        [Fact]
        public void HtmlToWord_ParagraphDirection_CssDirection() {
            string html = "<p style=\"direction: rtl\">Bravo</p><p style=\"direction:ltr\">Charlie</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var rtlParagraph = doc.Paragraphs.First(p => p.Text.Contains("Bravo"));
            var ltrParagraph = doc.Paragraphs.First(p => p.Text.Contains("Charlie"));

            Assert.True(rtlParagraph.BiDi);
            Assert.False(ltrParagraph.BiDi);
        }

        [Fact]
        public void HtmlToWord_BlockDirection_TextNode() {
            string html = "<div dir=\"rtl\">Delta</div>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var paragraph = doc.Paragraphs.First(p => p.Text.Contains("Delta"));
            Assert.True(paragraph.BiDi);
        }
    }
}
