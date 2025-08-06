using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_SpanStyles() {
            string html = "<p><span style=\"color:#ff0000;font-family:Arial;font-size:24px\">Styled</span></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();

            Assert.Equal("ff0000", run.ColorHex);
            Assert.Equal("Arial", run.FontFamily);
            Assert.Equal(24, run.FontSize);
        }
    }
}
