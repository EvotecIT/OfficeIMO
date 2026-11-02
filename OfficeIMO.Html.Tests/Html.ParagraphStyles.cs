using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_ParagraphStyles_ColorAndSize() {
            string html = "<p style=\"color:#ff0000;background-color:#ffff00;font-size:24px\">Styled</p>";

            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var paragraph = doc.Paragraphs[0];

            Assert.Equal(OfficeColor.Parse("FF0000"), paragraph.Color);
            Assert.Equal(18, paragraph.FontSize?.ToPoints());
            Assert.Equal(OfficeColor.Parse("FFFF00"), paragraph.ShadingFillColor);
        }

        [Fact]
        public void HtmlToWord_ParagraphStyles_NamedAndRgbColors() {
            string html = "<p style=\"color:red;background-color:rgb(0,255,255);font-size:20px\">Styled</p>";

            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var paragraph = doc.Paragraphs[0];

            Assert.Equal(OfficeColor.Parse("FF0000"), paragraph.Color);
            Assert.Equal(15, paragraph.FontSize?.ToPoints());
            Assert.Equal(OfficeColor.Parse("00FFFF"), paragraph.ShadingFillColor);
        }
    }
}
