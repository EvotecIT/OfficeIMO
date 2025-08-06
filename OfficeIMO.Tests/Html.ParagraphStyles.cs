using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_ParagraphStyles_ColorAndSize() {
            string html = "<p style=\"color:#ff0000;background-color:#ffff00;font-size:24px\">Styled</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];

            Assert.Equal("ff0000", paragraph.ColorHex);
            Assert.Equal(24, paragraph.FontSize);
            Assert.Equal(HighlightColorValues.Yellow, paragraph.Highlight);
        }

        [Fact]
        public void HtmlToWord_ParagraphStyles_NamedAndRgbColors() {
            string html = "<p style=\"color:red;background-color:rgb(0,255,255);font-size:20px\">Styled</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];

            Assert.Equal("ff0000", paragraph.ColorHex);
            Assert.Equal(20, paragraph.FontSize);
            Assert.Equal(HighlightColorValues.Cyan, paragraph.Highlight);
        }
    }
}
