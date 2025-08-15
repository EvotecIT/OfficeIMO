using System.Linq;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_LetterSpacing() {
            string html = "<p style=\"letter-spacing:2pt\">space</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();

            Assert.Equal(40, run.Spacing);
        }
    }
}
