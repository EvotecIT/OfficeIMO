using System.Linq;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_LetterSpacing() {
            string html = "<p style=\"letter-spacing:2pt\">spaced</p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var run = doc.Paragraphs.First();
            Assert.Equal(OfficeLength.Points(2), run.Spacing);
        }
    }
}
