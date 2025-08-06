using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_DivStyles_TextAlign() {
            string html = "<div style=\"text-align:center\"><p>Centered</p></div>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];

            Assert.Equal(JustificationValues.Center, paragraph.ParagraphAlignment);
        }

        [Fact]
        public void HtmlToWord_DivStyles_Margins() {
            string html = "<div style=\"margin-left:20pt;padding-left:10pt\"><p>Indented</p></div>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];

            Assert.Equal(30d, paragraph.IndentationBeforePoints);
        }
    }
}
